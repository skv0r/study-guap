%% ЛР2 по методичке: живой захват videoinput (устройство Color Device)
clc; close all; imaqreset;
base = fileparts(mfilename('fullpath'));
outDir = fullfile(base, 'screenshots');
if ~exist(outDir, 'dir'); mkdir(outDir); end
logDir = fullfile(base, 'capture');
if ~exist(logDir, 'dir'); mkdir(logDir); end

dll = fullfile(matlabroot, 'toolbox', 'imaq', 'imaqadaptors', 'kit', 'demo', 'win64', 'mwdemoimaq.dll');
if ~any(strcmp(imaqregister, dll))
    imaqregister(dll);
end
imaqreset;

%% Шаг 1-2. Адаптеры и ID
info = imaqhwinfo;
dev = imaqhwinfo('mwdemoimaq');
fid = fopen(fullfile(outDir, 'imaqhwinfo.txt'), 'w');
fprintf(fid, 'InstalledAdaptors: %s\nMATLAB: %s\nIAT: %s\n', ...
    strjoin(info.InstalledAdaptors, ', '), info.MATLABVersion, info.ToolboxVersion);
fprintf(fid, 'DeviceIDs: %s\n', mat2str([dev.DeviceIDs{:}]));
for i = 1:numel(dev.DeviceIDs)
    cam = imaqhwinfo('mwdemoimaq', dev.DeviceIDs{i});
    fprintf(fid, '\nID=%d %s default=%s\n', cam.DeviceID, cam.DeviceName, cam.DefaultFormat);
    if ~isempty(cam.SupportedFormats)
        fprintf(fid, '  %s\n', cam.SupportedFormats{:});
    end
end
fclose(fid);

%% Три режима
cfgs = {1, 'RGB_NTSC', 3; 1, 'S-Video', 3; 2, 'CCIR', 1};
rows = {};
for i = 1:size(cfgs, 1)
    vid = videoinput('mwdemoimaq', cfgs{i, 1}, cfgs{i, 2});
    src = getselectedsource(vid);
    res = vid.VideoResolution;
    fps = str2double(string(src.FrameRate));
    if isnan(fps), fps = 30; end
    B = cfgs{i, 3};
    D = res(1) * res(2) * B * fps / 1e6;
    snap = getsnapshot(vid);
    if size(snap, 3) == 1
        imwrite(snap, fullfile(outDir, sprintf('mode_%s.png', regexprep(cfgs{i, 2}, '\W', '_'))));
    else
        imwrite(snap, fullfile(outDir, sprintf('mode_%s.png', regexprep(cfgs{i, 2}, '\W', '_'))));
    end
    rows(end+1, :) = {cfgs{i, 2}, res(1), res(2), fps, B, D}; %#ok<SAGROW>
    delete(vid); imaqreset;
end
T = cell2table(rows, 'VariableNames', {'Format', 'W', 'H', 'FPS', 'BytesPerPixel', 'MBps'});
writetable(T, fullfile(outDir, 'stream.csv'));
disp(T);

%% Preview живого потока
vid = videoinput('mwdemoimaq', 1, 'RGB_NTSC');
vid.ReturnedColorSpace = 'rgb';
hImg = preview(vid);
pause(2.5);
hFig = ancestor(hImg, 'figure');
if isempty(hFig), hFig = gcf; end
fr = getframe(hFig);
imwrite(fr.cdata, fullfile(outDir, 'preview.png'));
imwrite(getsnapshot(vid), fullfile(outDir, 'snapshot_default.png'));
stoppreview(vid); closepreview(vid);

%% Свойства источника, два параметра
src = getselectedsource(vid);
props = get(src);
writecell(fieldnames(props), fullfile(outDir, 'source_prop_names.txt'));
fid = fopen(fullfile(outDir, 'source_props.txt'), 'w');
fn = fieldnames(props);
for k = 1:numel(fn)
    fprintf(fid, '%s = %s\n', fn{k}, string(props.(fn{k})));
end
fclose(fid);

src.Hue = 0.12;
src.Saturation = 92;
preview(vid); pause(2);
hImg = findobj(0, 'Type', 'image');
hFig = ancestor(hImg(1), 'figure');
imwrite(getframe(hFig).cdata, fullfile(outDir, 'prop_hue_sat_1.png'));
imwrite(getsnapshot(vid), fullfile(outDir, 'snapshot_hue_sat_high.png'));
stoppreview(vid); closepreview(vid);

src.Hue = 0.82;
src.Saturation = 12;
preview(vid); pause(2);
hImg = findobj(0, 'Type', 'image');
hFig = ancestor(hImg(1), 'figure');
imwrite(getframe(hFig).cdata, fullfile(outDir, 'prop_hue_sat_2.png'));
imwrite(getsnapshot(vid), fullfile(outDir, 'snapshot_hue_sat_low.png'));
stoppreview(vid); closepreview(vid);
src.Hue = 0.5;
src.Saturation = 50;
delete(vid); imaqreset;

%% Logging на диск 6 с (~180 кадров при 30 FPS)
vid = videoinput('mwdemoimaq', 1, 'RGB_NTSC');
vid.ReturnedColorSpace = 'rgb';
vid.LoggingMode = 'disk';
vid.FramesPerTrigger = 180;
vid.TriggerRepeat = 0;
aviPath = fullfile(logDir, 'log_6s.avi');
if exist(aviPath, 'file'), delete(aviPath); end
vw = VideoWriter(aviPath);
vid.DiskLogger = vw;
start(vid);
try
    wait(vid, 20, 'logging');
catch
    pause(8);
    if isrunning(vid), stop(vid); end
end
nWait = 0;
while vid.DiskLoggerFrameCount < vid.FramesAcquired && nWait < 50
    pause(0.1);
    nWait = nWait + 1;
end
fid = fopen(fullfile(outDir, 'logging.txt'), 'w');
fprintf(fid, 'LoggingMode=%s\nFramesPerTrigger=%d\nFramesAcquired=%d\nDiskLoggerFrameCount=%d\nFile=%s\n', ...
    vid.LoggingMode, vid.FramesPerTrigger, vid.FramesAcquired, vid.DiskLoggerFrameCount, aviPath);
fclose(fid);
if isrunning(vid), stop(vid); end
delete(vid); imaqreset;

%% Детект движения getdata, пороги 5/10/25, два разрешения
motion_live(1, 'RGB_NTSC', [5 10 25], outDir, 8);
motion_live(1, 'S-Video', 10, outDir, 6);

%% Очистка
imaqreset; close all;
fprintf('Методичка: захват завершён. AVI: %s\n', aviPath);

function motion_live(devId, fmt, thresholds, outDir, sec)
    if isscalar(thresholds), thresholds = thresholds; end
    for t = thresholds
        vid = videoinput('mwdemoimaq', devId, fmt);
        vid.ReturnedColorSpace = 'rgb';
        vid.TriggerRepeat = Inf;
        vid.FramesPerTrigger = 30;
        vid.FrameGrabInterval = 1;
        start(vid);
        saved = false;
        last = [];
        tic;
        while isrunning(vid) && toc < sec
            if vid.FramesAvailable >= 2
                data = getdata(vid, 2);
                I1 = data(:,:,:,1);
                I2 = data(:,:,:,2);
                if size(I1, 3) == 3
                    I1 = rgb2gray(I1);
                    I2 = rgb2gray(I2);
                end
                BW = imabsdiff(I1, I2) > t;
                if ~saved
                    imwrite(BW, fullfile(outDir, sprintf('motion_%s_th%d.png', regexprep(fmt, '\W', '_'), t)));
                    saved = true;
                end
                last = BW;
            elseif vid.FramesAvailable >= 1
                getdata(vid, 1);
            end
        end
        if ~saved && ~isempty(last)
            imwrite(last, fullfile(outDir, sprintf('motion_%s_th%d.png', regexprep(fmt, '\W', '_'), t)));
        end
        if isrunning(vid), stop(vid); end
        delete(vid); imaqreset; close all;
    end
end
