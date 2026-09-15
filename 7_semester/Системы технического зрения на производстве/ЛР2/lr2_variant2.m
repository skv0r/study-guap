%% ЛР2 вариант 2: захват, свойства, детект движения, поток, два разрешения
% Запускать после probe_setup.m. Скриншоты пишутся в ./screenshots

clc; close all;
outDir = fullfile(fileparts(mfilename('fullpath')), 'screenshots');
if ~exist(outDir, 'dir'); mkdir(outDir); end

imaqreset;
info = imaqhwinfo('winvideo');
devId = info.DeviceIDs{1};
cam = imaqhwinfo('winvideo', devId);
formats = cam.SupportedFormats(:);
fprintf('Камера: %s\nФорматов: %d\n', cam.DeviceName, numel(formats));
disp(formats);

% Три режима: мелкий, крупный несжатый, крупный MJPG если есть
pick = local_pick_formats(formats);
fid = fopen(fullfile(outDir, 'formats.txt'), 'w');
fprintf(fid, 'Device: %s\nDefault: %s\n\nAll formats:\n', cam.DeviceName, cam.DefaultFormat);
fprintf(fid, '%s\n', formats{:});
fprintf(fid, '\nChosen:\n');
fprintf(fid, '%s\n', pick{:});
fclose(fid);

%% Preview
vid = videoinput('winvideo', devId, pick{1});
vid.ReturnedColorSpace = 'rgb';
preview(vid);
pause(3);
fr = getframe(gcf);
imwrite(fr.cdata, fullfile(outDir, 'preview.png'));
stoppreview(vid); closepreview(vid);

%% Свойства источника
src = getselectedsource(vid);
props = get(src);
save(fullfile(outDir, 'source_props.mat'), 'props');
propNames = fieldnames(props);
writecell(propNames, fullfile(outDir, 'source_prop_names.txt'));

changed = {};
if isprop(src, 'Brightness')
    stoppreview(vid); closepreview(vid);
    src.Brightness = min(props.Brightness + 20, 255);
    preview(vid); pause(2);
    fr = getframe(gcf);
    imwrite(fr.cdata, fullfile(outDir, 'brightness.png'));
    changed{end+1} = sprintf('Brightness -> %s', num2str(src.Brightness));
    stoppreview(vid); closepreview(vid);
end
if isprop(src, 'Contrast')
    src.Contrast = min(props.Contrast + 10, 255);
    preview(vid); pause(2);
    fr = getframe(gcf);
    imwrite(fr.cdata, fullfile(outDir, 'contrast.png'));
    changed{end+1} = sprintf('Contrast -> %s', num2str(src.Contrast));
    stoppreview(vid); closepreview(vid);
end
if isprop(src, 'ExposureMode') && isprop(src, 'Exposure')
    try
        src.ExposureMode = 'manual';
        src.Exposure = props.Exposure;
        preview(vid); pause(2);
        fr = getframe(gcf);
        imwrite(fr.cdata, fullfile(outDir, 'exposure.png'));
        changed{end+1} = sprintf('ExposureMode=manual, Exposure=%s', num2str(src.Exposure));
        stoppreview(vid); closepreview(vid);
    catch e
        changed{end+1} = ['Exposure недоступен: ' e.message];
    end
end
writecell(changed(:), fullfile(outDir, 'changed_props.txt'));
delete(vid); clear vid; imaqreset;

%% Поток данных для трёх режимов
rows = {};
for i = 1:numel(pick)
    fmt = pick{i};
    vid = videoinput('winvideo', devId, fmt);
    src = getselectedsource(vid);
    res = vid.VideoResolution;
    W = res(1); H = res(2);
    fps = local_fps(src);
    B = local_bytes_per_pixel(fmt);
    D = W * H * B * fps / 1e6;
    rows(end+1, :) = {fmt, W, H, fps, B, D}; %#ok<SAGROW>
    delete(vid); clear vid; imaqreset;
end
T = cell2table(rows, 'VariableNames', {'Format', 'W', 'H', 'FPS', 'BytesPerPixel', 'MBps'});
disp(T);
writetable(T, fullfile(outDir, 'stream.csv'));

%% Детект движения + сравнение двух разрешений
motionFmts = local_motion_pair(pick, formats);
thresholds = [5, 10, 25];
for mi = 1:numel(motionFmts)
    fmt = motionFmts{mi};
    for ti = 1:numel(thresholds)
        local_motion_capture(devId, fmt, thresholds(ti), outDir, 40);
        imaqreset;
    end
end

%% Очистка
imaqreset; close all;
fprintf('Готово. Смотрите папку screenshots.\n');

function pick = local_pick_formats(formats)
    pick = {};
    small = formats(contains(formats, {'640x480', '320x240', '160x120'}));
    bigY = formats(contains(formats, {'1280x720', '1920x1080'}) & contains(formats, {'YUY', 'YUY2', 'YUV', 'Y800'}));
    bigJ = formats(contains(formats, {'1280x720', '1920x1080'}) & contains(formats, {'MJPG', 'JPEG'}));
    if ~isempty(small); pick{end+1} = small{1}; end
    if ~isempty(bigY); pick{end+1} = bigY{1}; end
    if ~isempty(bigJ); pick{end+1} = bigJ{1}; end
    rest = setdiff(formats, pick, 'stable');
    while numel(pick) < 3 && ~isempty(rest)
        pick{end+1} = rest{1};
        rest(1) = [];
    end
    pick = pick(1:min(3, numel(pick)));
end

function fps = local_fps(src)
    fps = 30;
    if isprop(src, 'FrameRate')
        fr = src.FrameRate;
        if ischar(fr) || isstring(fr)
            fps = str2double(fr);
        else
            fps = double(fr);
        end
        if isnan(fps) || fps <= 0
            fps = 30;
        end
    end
end

function B = local_bytes_per_pixel(fmt)
    u = upper(fmt);
    if contains(u, {'YUY', 'YUV', 'YUY2'})
        B = 2;
    elseif contains(u, {'MJPG', 'JPEG'})
        B = 3; % эквивалент RGB после декодирования
    elseif contains(u, {'RGB', 'ARGB'})
        B = 3;
    else
        B = 3;
    end
end

function pair = local_motion_pair(pick, formats)
    mjpg = formats(contains(formats, 'MJPG') & contains(formats, {'640x480', '1280x720'}));
    if numel(mjpg) >= 2
        pair = mjpg(1:2);
        return;
    end
    pair = pick(1:min(2, numel(pick)));
end

function local_motion_capture(devId, fmt, threshold, outDir, nFrames)
    vid = videoinput('winvideo', devId, fmt);
    vid.ReturnedColorSpace = 'rgb';
    vid.TriggerRepeat = Inf;
    vid.FramesPerTrigger = 20;
    vid.FrameGrabInterval = 1;
    start(vid);
    last = [];
    saved = false;
    got = 0;
    tEnd = tic;
    while isrunning(vid) && toc(tEnd) < 8
        if vid.FramesAvailable >= 1
            data = getdata(vid, 1);
            I = rgb2gray(data(:,:,:,1));
            if ~isempty(last)
                BW = imabsdiff(I, last) > threshold;
                got = got + 1;
                if ~saved && any(BW(:))
                    imwrite(BW, fullfile(outDir, sprintf('motion_%s_th%d.png', local_slug(fmt), threshold)));
                    saved = true;
                end
            end
            last = I;
        end
        if got >= nFrames
            break;
        end
    end
    if ~saved && ~isempty(last)
        BW = last > 0;
        imwrite(BW, fullfile(outDir, sprintf('motion_%s_th%d.png', local_slug(fmt), threshold)));
    end
    stop(vid); delete(vid); clear vid;
end

function s = local_slug(fmt)
    s = regexprep(fmt, '[^A-Za-z0-9]+', '_');
end
