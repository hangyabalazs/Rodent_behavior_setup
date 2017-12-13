function TeensyCalibration
%TEENSYCALIBRATION   Sound calibration using Bpod and Teensy sound server.
%   TEENSYCALIBRATION generates a sound calibration lookup table to allow
%   the Bpod (behavior control device) to play calibrated sounds of certain
%   sound pressure levels at selected frequencies.
%
% NOTE: this code has been used on a Bpod State Machine r0.5. It is not 
% guaranteed to work on r0.8 or r2 generation without modifications in the 
% connection and command lines, although the core mechanism remains the same
%
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%                               INSTRUCTIONS                              %
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%
%                               What you need:
% 
% - Working Bpod State Machine r0.5 (out of production, PCB available only) 
% - TrueRTA software (https://www.trueaudio.com/rta_down.htm) 
% Download and install the free version. If you already posess another
% similar software that reads dB SPL from a microphone you can use it
% instead (edit the code accordingly).
% - Calibrated microphone (set it as input device in TrueRTA's Audio I/O -> 
% Audio Device Selection menu, Audio Input Devide Selection popup menu)
% - Teensy with Audioboard(https://sites.google.com/site/bpoddocumentation/
% bpod-user-guide/function-reference/teensysoundserver)
%
%                               Instructions:
%
% This protocol will create calibrated pure tones from 1kHz to 21kHz.
% At the beginning you will be prompted to specify the frequency resolution:
%     - 1kHz to produce 1kHz, 2kHz, 3kHz, etc.. sinewaves
%     - 0.5Khz to produce 1kHz, 1.5kHz,  2kHz, 2.5kHz, 3kHz, etc.. sinewaves
%     - 0.25kHz  to produce 1kHz, 1.25kHz, 1.5kHz, 1.75kHz, 2kHz, 2.25kHz,
%       2.5kHz, etc.. sinewaves
% The program will generate the corresponding audio files, create an Excel
% .xlsx file and open TrueRTA. The tones will be played using the Bpod
% interface. You have to read the dB SPL values of the sounds played in
% TrueRTA and enter them in the appropriate field of the Excel table. Each
% dB SPL should be written in the cell next to the corresponding frequency
% listed in the first column of the table.
%
% First, LED2 turns on and after 1 second a tone is played. After that, the
% user should click on Port3 to proceed with the next tone or Port1 to play
% the tone again. When done, save and close the xlsx. This procedure is
% repeated twice to have an accurate calibration. Finally, a similar
% procedure is applied to create a calbrated white noise track. Calibration
% data is be saved in the TeensyCaldata.mat file which can be used for
% future experiments. Dialogue boxes guides the user through the process.
%
%                                 Note: 
%
% Since Teensy lacks a specific ID and is recognized during the Teensy
% Sound Server initialization as a generic USB device, this can create
% conflict if other similar devices are plugged in.
% 
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Nicola Solari, 2016
% solari.nicola@koki.mta.hu
% Lendulet Laboratory of Systems Neuroscience (hangyalab.koki.hu)
% Institute of Experimental Medicine, Hungarian Academy of Sceinces

% Load Bpod object
global BpodSystem % Allows access to Bpod device parameters from this function

% Cleanup serial ports
delete(instrfindall)

% Initialize Teensy Sound Server
[Status RawString] = system('wmic path Win32_SerialPort Where "Caption LIKE ''%USB%''" Get PNPDeviceID');
 Specific_ID = RawString(end-12:end);
Part1 = 'wmic path Win32_SerialPort Where "PNPDeviceID LIKE ''%';
Part2 = Specific_ID;
Part3 = '%''" Get DeviceID';
magicformula = strcat (Part1,Part2,Part3);
[Status RawString] = system(magicformula);PortLocations = strfind(RawString, 'COM');
TeensyPorts = cell(1,100);
nPorts = length(PortLocations);
for x = 1:nPorts
    Clip = RawString(PortLocations(x):PortLocations(x)+6);
    TeensyPorts{x} = Clip(1:find(Clip == 32,1, 'first')-1);
end
TeensyPort = TeensyPorts(1:nPorts);
TeensySoundServer('init',TeensyPort); % Opens Teensy Server at the COM port engaged by the device

% Lauch TrueRTA
winopen('C:\Program Files (x86)\TrueRTA_3\TrueRTA.exe') % Opens TrueRTA (please, verify and eventually change the path if installed elsewhere)

% Sound calibration
choice = questdlg({'This procedure generates calibrated pure tones from 1kHz to 21kHz.';'Select the sampling interval to apply in this range.';'1kHz (generates 21 tracks: 1kHz, 2Khz, 3Kz, 4kHz, etc..)';...
    '0.5kHz (generates 41 tracks: 1kHz, 1.5kHz, 2Khz, 2.5Kz, etc..)';'0.25kHz (generates 81 tracks: 1kHz, 1.25kHz, 1.5kHz, 1.75kHz, etc..)'},'FREQUENCY RESOLUTION','1kHz','0.5kHz','0.25kHz','0.25kHz');
xlsx = 'Calibration.xlsx';
switch choice
    case '1kHz' % Consider each frequency from 1kHz to 21kHz with a 1kHz interval
        Span = 1000; NumHz = 21;
        FreqY = (1:1:21)';
        xlswrite(xlsx,FreqY);
        RangeY = 'B2:B22';
    case '0.5kHz' % Consider each frequency from 1kHz to 21kHz with a 0.5kHz interval
        Span = 500; NumHz = 41;
        FreqY = (1:0.5:21)';
        xlswrite(xlsx,FreqY);
        RangeY = 'B2:B42';
    case '0.25kHz' % Consider each frequency from 1kHz to 21kHz with a 0.25kHz interval
        Span = 250; NumHz = 81;
        FreqY = (1:0.25:21)';
        xlswrite(xlsx,FreqY);
        RangeY = 'B2:B82';
end
TrackY = (1:NumHz)';
AmplY = repelem (0.3, NumHz)';
Header={'kHz Frequency','dB SPL','Amplitude','Track Number'};     
xlswrite(xlsx,Header,1,'A1'); % Prepare xlsx table
xlswrite(xlsx,FreqY,1,'A2');
xlswrite(xlsx,TrackY,1,'D2'); 
xlswrite(xlsx,AmplY,1,'C2');
Pre = NaN(1, NumHz)'; % Preallocations
[Ampl, CalAmpl, Ampl1, Ampl2] = deal(Pre);
loader(NumHz,Span,TeensyPort); % Load sinewaves with a preset amplitude of 0.3 on Teensy SD 
winopen(xlsx);
Test(NumHz); % Plays the loaded tracks with Bpod interface
f = warndlg('Save and close Calibration.xlsx before venturing forth', 'STOP');
waitfor(f);
x = inputdlg('Enter the desired SPL value','Calibration Target',1);   % Choose the dB SPL common to all your tracks
Tg = str2num(x{:});
SPL = xlsread(xlsx,RangeY);
CalAmpl = AmplAdjst(SPL,Tg,AmplY); % Adjust sinewaves amplitude to have the desired dB SPL
SineRewrite (NumHz,Span,CalAmpl,TeensyPort); % Rewrite the sinewaves with ne newly formed amplitude
xlswrite(xlsx,CalAmpl,1,'C2');
winopen (xlsx);
Test(NumHz); % Plays the loaded tracks with Bpod Interface
f = warndlg('Save and close Calibration.xlsx before venturing forth', 'STOP');
waitfor(f);
SPL = xlsread(xlsx,RangeY);
Ampl1 = CalAmpl;
Ampl2 = AmplAdjst(SPL,Tg,Ampl1); % Adjust sinewaves amplitude to have a more precise dB SPL
SineRewrite (NumHz,Span, Ampl2,TeensyPort); % Rewrite the sinewaves with ne newly formed amplitude
xlswrite(xlsx,Ampl2,1,'C2');
d = dialog('Position',[500 600 250 95],'Name','Wait..');
txt = uicontrol('Parent',d,'Style','text','Position',[20 10 210 60],'String',{'The calibration process is completed.';'Please check the refinement.'});
waitfor (d);
Test(NumHz); % Plays the loaded tracks with Bpod Interface

% Whitenoise calibration
choice = questdlg({'Your calibrated tracks have been uploaded';'Would you like also to add a white noise track?'},'Almost done...','No, thanks','Yes, please','Yes, please')
switch choice
    case 'Yes, please' % Start the creation of a calibrated whitenoise track
        m = dialog('WindowStyle', 'normal','Position',[500 600 250 95],'Name','Wait..');
        txt = uicontrol('Parent',m,'Style','text','Position',[20 10 210 60],'String',{'Loading whitenoise.';'Please wait.'});
        if NumHz == 21; % Prepare xlsx table references accordingly
            RangeY = 'A23:A28';
            SPLY = 'B24:B28';
            nnloc = 'D23';
            Amploc = 'C23';
            Tgloc = 'B23';
        elseif NumHz == 41;
            RangeY = 'A43:A48'
            SPLY = 'B44:B48';
            nnloc = 'D43';
            Amploc = 'C43';
            Tgloc = 'B43';
        else
            RangeY = 'A83:A88'
            SPLY = 'B84:B88';
            nnloc = 'D83';
            Amploc = 'C83';
            Tgloc = 'B83';
        end
        FreqY = {99 '1kHzOct' '2kHzOct' '5kHzOct' '10kHzOct' '20kHzOct'}';
        xlswrite(xlsx,FreqY,RangeY); % Extend xlsx table
        xlswrite(xlsx,99,1,nnloc);
        White = randn(1,6*44100)*0.3; % Create the whitenoise
        TeensySoundServer('loadwaveform', 99, White);    %Load the whitenoise track as Track number 99
        close(m);
        f = warndlg('Play the whitenoise (identified as frquency 99) track with Bpod interface. Please write the dB SPL value observed at 1kHz, 2kHz, 5kHz, 10kHz and 20kHz octave near their newly appeared references','STOP');
        waitfor(f);
        winopen (xlsx);
        WhitePlay; % Plays the whitenoisek with Bpod Interface
        f = warndlg('Save and close Calibration.xlsx before venturing forth.', 'STOP');
        waitfor(f);
        m = dialog('WindowStyle', 'normal','Position',[500 600 250 95],'Name','Wait..');
        txt = uicontrol('Parent',m,'Style','text','Position',[20 10 210 60],'String',{'Calibrating whitenoise.';'Please wait.'});
        SPL = xlsread(xlsx,SPLY);
        Ampl = repelem (0.3, 5)';
        CalAmpl = AmplAdjst(SPL,Tg,Ampl); % Adjust reported whitenoise components amplitude
        WhiteAmp = mean(CalAmpl);
        xlswrite(xlsx,WhiteAmp,1,Amploc);
        xlswrite(xlsx,Tg,1,Tgloc);
        WhiteRewrite (WhiteAmp,TeensyPort); % Rewrites whitenoise track
        close(m);
        m = dialog('WindowStyle', 'modal','Position',[500 600 250 95],'Name','Wait..');
        txt = uicontrol('Parent',m,'Style','text','Position',[20 10 210 60],'String',{'Whitenoise has been calibrated.';'Please check it.'});
        waitfor(m);
        WhitePlay;
    case 'No, thanks'
        msgbox('Well, then we have finished.')
end

% Calibration Struct creation
if NumHz == 21;
    Frequency = 'A2:A23';
    Volume = 'B2:B23';
    Amplitude = 'C2:C23';
elseif NumHz == 41;
    Frequency = 'A2:A43';
    Volume = 'B2:B43';
    Amplitude = 'C2:C43';
else
    Frequency = 'A2:A83';
    Volume = 'B2:B83';
    Amplitude = 'C2:C83';
end
A = xlsread ('Calibration.xlsx', Frequency);
B = xlsread ('Calibration.xlsx', Volume);
C = xlsread ('Calibration.xlsx', Amplitude);
TeensyCalData = struct('Frequency',A,'SPL',B,'Amplitude',C); % Create a struct with calibration data
save('TeensyCalData', 'TeensyCalData*');

msgbox('Mission Complete. Godspeed.')
TeensySoundServer('close'); % Closes Teensy Sound Server


% -------------------------------------------------------------------------
function loader(NumHz,Span,TeensyPort) % Load the standard sinewaves with the arbitrary pre-set amplitude of 0.3

SamplingRate = 44100;
SoundDuration = 2;
Pre = NaN(1, NumHz);  % Preallocation
Pre2 = num2cell (Pre);   
[sinewave] = deal(Pre2);
load.a = repelem (0.3, NumHz);
load.b = 1000:Span:(21000);
h = waitbar(0,{'This is quick'});
for n= 1:NumHz;
    sinewave (:,n) = {load.a(n).*sin(2*pi*load.b(n)/SamplingRate.*(0:SamplingRate*SoundDuration))'}; % Creates the sinewaves
    waitbar(n / NumHz)
end
close(h)
delete(instrfindall);
TeensySoundServer('init',TeensyPort'); % Initialize Teensy Sound Server 
h = waitbar(0,{'Meanwhile you can minimize Matlab window and read the instructions:';'LED2 will light and a sound will be played shortly after';...
    'Click Port3 to move to the next trial';'Click Port1 to replay the current one'});
set(h,'Name','LOADING SINEWAVES..');
for  n = 1:NumHz;
    TeensySoundServer ('loadwaveform', n, sinewave{n}); % Loads the sinewaves
    waitbar(n / NumHz)
end
close(h)

% -------------------------------------------------------------------------
function [CalAmpl] = AmplAdjst(SPL,Tg,Ampl) % Calculate the new proper sinewave amplitude

y = SPL - Tg;
b =  20 * log10(Ampl) - y;
c = b / 20;
CalAmpl = 10 .^ c;

% -------------------------------------------------------------------------
function SineRewrite(NumHz,Span,CalAmpl,TeensyPort)  % Load new sinewaves, it takes about 30sec for 21 tracks, 1min for 41 and 2min for 81

SamplingRate = 44100;
SoundDuration = 2;
Pre = NaN(1, NumHz); % Preallocation
Pre2 = num2cell (Pre);
[sinewave] = deal(Pre2);
range = 1000:Span:(NumHz*1000);
sound.frequency = range;
sound.ampl = CalAmpl;
for n=1:NumHz;
    sinewave (:,n) = {sound.ampl(n).*sin(2*pi*sound.frequency(n)/SamplingRate.*(0:SamplingRate*SoundDuration))'}; % Creates the sinewaves
end
delete(instrfindall);
TeensySoundServer('init',TeensyPort); % Initialize Teensy Sound Server
h = waitbar(0,{'Wait, calibration in progress'});
set(h,'Name','REWRITING');
for  n = 1:NumHz;
    TeensySoundServer ('loadwaveform', n, sinewave{n}) % Loads the sinewaves
    waitbar(n / NumHz)
end
close(h)

% -------------------------------------------------------------------------
function Test(NumHz)  % Plays the tracks loaded on the Teensy SD card (see instructions for the dynamic)

global BpodSystem
for currentTrial = 1:NumHz;
    sma = NewStateMatrix(); % Assemble state matrix
    k = currentTrial;
    sma = AddState(sma, 'Name', 'Beginning', ...   % A shord delay to prepare yourself
        'Timer', 1,...
        'StateChangeConditions', {'Tup', 'PlaySine'},...
        'OutputActions', {'PWM2', 255});
    sma = AddState(sma, 'Name', 'PlaySine', ...        % Tone will be played
        'Timer', 1,...
        'StateChangeConditions', {'Tup', 'Crossroad'},...
        'OutputActions', {'PWM2', 255, 'Serial1Code', k});
    sma = AddState(sma, 'Name', 'Crossroad', ...       % Click on Port3 to go to the next trial, Port1 to repeat
        'Timer', 0,...
        'StateChangeConditions', {'Port3In', 'exit','Port1In', 'PlaySine'},...
        'OutputActions', {'PWM1', 255, 'PWM3', 255, 'PWM2', 0});
    SendStateMatrix(sma);
    RawEvents = RunStateMatrix; % Send and run state matrix
    HandlePauseCondition; % Checks to see if the protocol is paused. If so, waits until user resumes.
    if BpodSystem.BeingUsed == 0
        return
    end
end

% -------------------------------------------------------------------------
function WhitePlay  % Plays the whitenoise loaded on the Teensy SD card

global BpodSystem
for currentTrial = 1
    sma = NewStateMatrix(); % Assemble state matrix
    sma = AddState(sma, 'Name', 'Beginning', ...   % W A shord delay to prepare yourself
        'Timer', 1,...
        'StateChangeConditions', {'Tup', 'PlaySine'},...
        'OutputActions', {'PWM2', 255});
    sma = AddState(sma, 'Name', 'PlaySine', ...        % Tone will be played
        'Timer', 6,...
        'StateChangeConditions', {'Tup', 'Crossroad'},...
        'OutputActions', {'PWM2', 255, 'Serial1Code', 99});
    sma = AddState(sma, 'Name', 'Crossroad', ...       % Click on Port3 to end the trial, Port1 to repeat
        'Timer', 0,...
        'StateChangeConditions', {'Port3In', 'exit','Port1In', 'PlaySine'},...
        'OutputActions', {'PWM1', 255, 'PWM3', 255});
    SendStateMatrix(sma);
    RawEvents = RunStateMatrix; % Send and run state matrix
    HandlePauseCondition; % Checks to see if the protocol is paused. If so, waits until user resumes.
    if BpodSystem.BeingUsed == 0
        return
    end
end

% -------------------------------------------------------------------------
function WhiteRewrite(WhiteAmp, TeensyPort)  % Load new whitenoise track

White = randn(1,6*44100)*WhiteAmp;
delete(instrfindall);
TeensySoundServer('init',TeensyPort); % Initialize Teensy Sound Server
TeensySoundServer ('loadwaveform', 99, White); % Load whitenoise