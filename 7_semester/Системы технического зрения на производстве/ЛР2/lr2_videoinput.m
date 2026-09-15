%% ЛР2: захват через videoinput (mwdemoimaq Color Device)
clc; close all; imaqreset;
base = fileparts(mfilename('fullpath'));
outDir = fullfile(base, 'screenshots');
logDir = fullfile(base, 'capture');
if ~exist(outDir, 'dir'); mkdir(outDir); end
if ~exist(logDir, 'dir'); mkdir(logDir); end
loadCsv = fullfile(outDir, 'load.csv');
if exist(loadCsv, 'file'), delete(loadCsv); end

dll = fullfile(matlabroot, 'toolbox', 'imaq', 'imaqadaptors', 'kit', 'demo', 'win64', 'mwdemoimaq.dll');
if ~any(strcmp(imaqregister, dll))
    imaqregister(dll);
end
imaqreset;

%% Шаг 1–2. Адаптеры, ID, форматы
info = imaqhwinfo;
dev = imaqhwinfo('mwdemoimaq');
fid = fopen(fullfile(outDir, 'imaqhwinfo.txt'), 'w');
fprintf(fid, 'InstalledAdaptors: %s\nMATLAB: %s\nToolbox: %s %s\n', ...
    strjoin(info.InstalledAdaptors, ', '), info.MATLABVersion, ...
    info.ToolboxName, info.ToolboxVersion);
fprintf(fid, 'DeviceIDs: %s\n', mat2str([dev.DeviceIDs{:}]));
for i = 1:numel(dev.DeviceIDs)
    cam = imaqhwinfo('mwdemoimaq', dev.DeviceIDs{i});
    fprintf(fid, '\nID=%d %s default=%s DeviceFileSupported=%d\n', ...
        cam.DeviceID, cam.DeviceName, cam.DefaultFormat, cam.DeviceFileSupported);
    if ~isempty(cam.SupportedFormats)
        fprintf(fid, '  %s\n', cam.SupportedFormats{:});
    end
end
fclose(fid);
write_text_fig(fileread(fullfile(outDir, 'imaqhwinfo.txt')), ...
    fullfile(outDir, 'imaqhwinfo.png'), 'imaqhwinfo(''mwdemoimaq'')');

%% Три режима: snapshot, разрешение, FPS, поток
cfgs = {1, 'RGB_NTSC', 3; 1, 'S-Video', 3; 2, 'CCIR', 1};
rows = {};
for i = 1:size(cfgs, 1)
    vid = videoinput('mwdemoimaq', cfgs{i, 1}, cfgs{i, 2});
    src = getselectedsource(vid);
    res = vid.VideoResolution;
    fps = str2double(src.FrameRate);
    if isnan(fps), fps = 30; end
    B = cfgs{i, 3};
    D = res(1) * res(2) * B * fps / 1e6;
    snap = getsnapshot(vid);
    fname = sprintf('mode_%s.png', regexprep(cfgs{i, 2}, '\W', '_'));
    imwrite(snap, fullfile(outDir, fname));
    rows(end+1, :) = {cfgs{i, 2}, res(1), res(2), fps, B, D}; %#ok<SAGROW>
    delete(vid); imaqreset;
end
T = cell2table(rows, 'VariableNames', {'Format', 'W', 'H', 'FPS', 'BytesPerPixel', 'MBps'});
writetable(T, fullfile(outDir, 'stream.csv'));
disp(T);

%% Preview живого потока RGB_NTSC
vid = videoinput('mwdemoimaq', 1, 'RGB_NTSC');
vid.ReturnedColorSpace = 'rgb';
src = getselectedsource(vid);
hImg = preview(vid);
pause(2.5);
save_preview_frame(hImg, fullfile(outDir, 'preview.png'));
imwrite(getsnapshot(vid), fullfile(outDir, 'snapshot_default.png'));
stoppreview(vid); closepreview(vid);

%% Свойства источника: Hue, Saturation (ReadOnly whileRunning)
fid = fopen(fullfile(outDir, 'source_props.txt'), 'w');
dump_src(fid, src, 'default');
src.Hue = 0.12;
src.Saturation = 92;
dump_src(fid, src, 'high');
hImg = preview(vid);
pause(2.0);
save_preview_frame(hImg, fullfile(outDir, 'prop_hue_sat_1.png'));
imwrite(getsnapshot(vid), fullfile(outDir, 'snapshot_hue_sat_high.png'));
stoppreview(vid); closepreview(vid);

src.Hue = 0.82;
src.Saturation = 12;
dump_src(fid, src, 'low');
hImg = preview(vid);
pause(2.0);
save_preview_frame(hImg, fullfile(outDir, 'prop_hue_sat_2.png'));
imwrite(getsnapshot(vid), fullfile(outDir, 'snapshot_hue_sat_low.png'));
stoppreview(vid); closepreview(vid);
fclose(fid);
src.Hue = 0.5;
src.Saturation = 50;
write_text_fig(fileread(fullfile(outDir, 'source_props.txt')), ...
    fullfile(outDir, 'source_props.png'), 'getselectedsource: Hue / Saturation');

fid = fopen(fullfile(outDir, 'camera_info.txt'), 'w');
fprintf(fid, ['Adaptor=mwdemoimaq\nDeviceName=Color Device\nDeviceID=1\n', ...
    'Format=RGB_NTSC\nVideoResolution=%s\nReturnedColorSpace=%s\n', ...
    'FrameRate=%s\nHue=%g\nSaturation=%g\n'], ...
    mat2str(vid.VideoResolution), vid.ReturnedColorSpace, src.FrameRate, src.Hue, src.Saturation);
fclose(fid);
delete(vid); imaqreset;

%% LoggingMode=disk, 6 с (180 кадров при 30 FPS)
vid = videoinput('mwdemoimaq', 1, 'RGB_NTSC');
vid.ReturnedColorSpace = 'rgb';
vid.LoggingMode = 'disk';
vid.FramesPerTrigger = 180;
vid.TriggerRepeat = 0;
aviPath = fullfile(logDir, 'log_6s.avi');
if exist(aviPath, 'file'), delete(aviPath); end
vw = VideoWriter(aviPath, 'Motion JPEG AVI');
vw.FrameRate = 30;
vid.DiskLogger = vw;
start(vid);
t0 = tic;
while islogging(vid) && toc(t0) < 25
    pause(0.2);
end
if isrunning(vid), stop(vid); end
nWait = 0;
while vid.DiskLoggerFrameCount < vid.FramesAcquired && nWait < 80
    pause(0.1);
    nWait = nWait + 1;
end
fid = fopen(fullfile(outDir, 'logging.txt'), 'w');
fprintf(fid, ['LoggingMode=%s\nFramesPerTrigger=%d\nFramesAcquired=%d\n', ...
    'DiskLoggerFrameCount=%d\nFile=%s\nDuration_s=%.1f\nFPS=30\n'], ...
    vid.LoggingMode, vid.FramesPerTrigger, vid.FramesAcquired, ...
    vid.DiskLoggerFrameCount, aviPath, vid.FramesAcquired / 30);
fclose(fid);
write_text_fig(fileread(fullfile(outDir, 'logging.txt')), ...
    fullfile(outDir, 'logging.png'), 'LoggingMode = disk');
if isrunning(vid), stop(vid); end
delete(vid); imaqreset;

%% Детектор движения: соседние кадры getdata, пороги 5/10/25
capture_motion(1, 'RGB_NTSC', [5 10 25], outDir);
capture_motion(1, 'S-Video', 10, outDir);

%% Живой цикл ~4 с (скрин «в реальном времени»)
vid = videoinput('mwdemoimaq', 1, 'RGB_NTSC');
vid.ReturnedColorSpace = 'rgb';
vid.TriggerRepeat = Inf;
vid.FramesPerTrigger = 30;
fig = figure('Color', 'w', 'Name', 'Области движения');
start(vid);
t0 = tic;
while isrunning(vid) && toc(t0) < 4
    if vid.FramesAvailable >= 2
        data = getdata(vid, 2);
        BW = imabsdiff(rgb2gray(data(:,:,:,1)), rgb2gray(data(:,:,:,2))) > 10;
        imshow(BW, 'Parent', gca);
        title(sprintf('Порог = 10, FramesAcquired = %d', vid.FramesAcquired));
        drawnow;
    end
end
if isrunning(vid), stop(vid); end
imwrite(getframe(fig).cdata, fullfile(outDir, 'motion_live.png'));
delete(vid); close(fig); imaqreset;

%% Нагрузка: два разрешения, одни и те же 20 пар кадров
bench_res(1, 'RGB_NTSC', 20, outDir);
bench_res(1, 'S-Video', 20, outDir);

%% Очистка
imaqreset; close all;
fprintf('Готово. AVI %s\n', aviPath);

function capture_motion(devId, fmt, thresholds, outDir)
    if isscalar(thresholds), thresholds = thresholds; end
    vid = videoinput('mwdemoimaq', devId, fmt);
    vid.ReturnedColorSpace = 'rgb';
    vid.FramesPerTrigger = 8;
    start(vid);
    wait(vid, 12, 'logging');
    n = min(2, vid.FramesAvailable);
    data = getdata(vid, n);
    if isrunning(vid), stop(vid); end
    delete(vid); imaqreset;
    if size(data, 4) < 2
        warning('Мало кадров для %s', fmt);
        return;
    end
    I1 = data(:,:,:,1); I2 = data(:,:,:,2);
    if size(I1, 3) == 3
        I1 = rgb2gray(I1); I2 = rgb2gray(I2);
    end
    for t = thresholds
        BW = imabsdiff(I1, I2) > t;
        imwrite(BW, fullfile(outDir, sprintf('motion_%s_th%d.png', regexprep(fmt, '\W', '_'), t)));
    end
end

function bench_res(devId, fmt, nPairs, outDir)
    vid = videoinput('mwdemoimaq', devId, fmt);
    vid.ReturnedColorSpace = 'rgb';
    vid.TriggerRepeat = Inf;
    vid.FramesPerTrigger = 40;
    start(vid);
    dt = zeros(nPairs, 1);
    got = 0;
    t0 = tic;
    while got < nPairs && toc(t0) < 20
        if vid.FramesAvailable >= 2
            tic;
            data = getdata(vid, 2);
            a = data(:,:,:,1); b = data(:,:,:,2);
            if size(a, 3) == 3
                a = rgb2gray(a); b = rgb2gray(b);
            end
            BW = imabsdiff(a, b) > 10; %#ok<NASGU>
            got = got + 1;
            dt(got) = toc;
        end
    end
    acquired = vid.FramesAcquired;
    if isrunning(vid), stop(vid); end
    delete(vid); imaqreset;
    dt = dt(1:max(got, 1));
    if got < 1
        persist_load(outDir, fmt, 0, acquired, NaN, NaN);
    else
        persist_load(outDir, fmt, got, acquired, mean(dt)*1000, max(dt)*1000);
    end
end

function persist_load(outDir, fmt, pairs, acquired, meanMs, maxMs)
    csvFile = fullfile(outDir, 'load.csv');
    needHeader = (exist(csvFile, 'file') ~= 2);
    fid = fopen(csvFile, 'a');
    if needHeader
        fprintf(fid, 'Format,Pairs,FramesAcquired,MeanCycle_ms,MaxCycle_ms\n');
    end
    fprintf(fid, '%s,%d,%d,%.2f,%.2f\n', fmt, pairs, acquired, meanMs, maxMs);
    fclose(fid);
end

function dump_src(fid, src, tag)
    fprintf(fid, '[%s] FrameRate=%s Hue=%g Saturation=%g SyncInput=%s\n', ...
        tag, src.FrameRate, src.Hue, src.Saturation, src.SyncInput);
end

function save_preview_frame(hImg, path)
    hFig = ancestor(hImg, 'figure');
    if isempty(hFig), hFig = gcf; end
    drawnow;
    imwrite(getframe(hFig).cdata, path);
end

function write_text_fig(txt, path, titleStr)
    fig = figure('Color', 'w', 'Position', [80 80 760 420], 'MenuBar', 'none');
    axis off;
    title(titleStr, 'FontSize', 12, 'FontWeight', 'bold', 'Interpreter', 'none');
    text(0.02, 0.92, txt, 'FontName', 'Consolas', 'FontSize', 9, ...
        'VerticalAlignment', 'top', 'Interpreter', 'none', 'Units', 'normalized');
    drawnow;
    imwrite(getframe(fig).cdata, path);
    close(fig);
end
