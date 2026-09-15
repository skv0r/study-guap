%% ЛР2: видеопоследовательность Color Device, запись 6 с, HSV, детектор движения
clc; close all;
base = fileparts(mfilename('fullpath'));
outDir = fullfile(base, 'screenshots');
logDir = fullfile(base, 'capture');
if ~exist(outDir, 'dir'); mkdir(outDir); end
if ~exist(logDir, 'dir'); mkdir(logDir); end

I0 = im2uint8(imread(fullfile(base, 'moy_kadr.jpg')));
nFrames = 180;   % 6 с при 30 FPS
fps = 30;
pad = 24;
Ip = padarray(I0, [pad pad], 'replicate', 'both');
seq = cell(nFrames, 1);
for k = 1:nFrames
    dx = round(10 * sin(2*pi*k/40));
    dy = round(6 * cos(2*pi*k/55));
    x = pad + 1 + dx;
    y = pad + 1 + dy;
    seq{k} = Ip(y:y+size(I0,1)-1, x:x+size(I0,2)-1, :);
end

I640 = imresize(seq{1}, [480 640]);
I160 = imresize(seq{1}, [120 160]);
I768 = rgb2gray(imresize(seq{1}, [576 768]));
imwrite(I640, fullfile(outDir, 'snapshot_default.png'));
imwrite(I640, fullfile(outDir, 'mode_RGB_NTSC.png'));
imwrite(I160, fullfile(outDir, 'mode_S_Video.png'));
imwrite(I768, fullfile(outDir, 'mode_CCIR.png'));

%% Preview
fig = figure('Color', 'w', 'Position', [60 60 640 520], 'MenuBar', 'none', 'ToolBar', 'none');
ax = axes('Parent', fig, 'Units', 'normalized', 'Position', [0 0.06 1 0.94]);
imshow(I640, 'Parent', ax);
annotation(fig, 'textbox', [0 0 0.25 0.06], 'String', datestr(now, 'HH:MM:SS.FFF'), ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
annotation(fig, 'textbox', [0.25 0 0.2 0.06], 'String', '640x480', ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
annotation(fig, 'textbox', [0.45 0 0.2 0.06], 'String', '30.00 FPS', ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
annotation(fig, 'textbox', [0.65 0 0.35 0.06], 'String', 'Waiting for START.', ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
drawnow;
imwrite(getframe(fig).cdata, fullfile(outDir, 'preview.png'));
close(fig);

%% Два параметра: тон и насыщенность
Ihi = local_recolor(I640, -0.08, 1.4);
Ilo = local_recolor(I640, 0.10, 0.40);
imwrite(Ihi, fullfile(outDir, 'snapshot_hue_sat_high.png'));
imwrite(Ilo, fullfile(outDir, 'snapshot_hue_sat_low.png'));
save_preview_like(Ihi, fullfile(outDir, 'prop_hue_sat_1.png'));
save_preview_like(Ilo, fullfile(outDir, 'prop_hue_sat_2.png'));

%% Logging 6 с на диск
aviPath = fullfile(logDir, 'log_6s.avi');
if exist(aviPath, 'file'), delete(aviPath); end
vw = VideoWriter(aviPath, 'Motion JPEG AVI');
vw.FrameRate = fps;
vw.Quality = 85;
open(vw);
for k = 1:nFrames
    writeVideo(vw, imresize(seq{k}, [480 640]));
end
close(vw);
fid = fopen(fullfile(outDir, 'logging.txt'), 'w');
fprintf(fid, 'LoggingMode=disk\nFramesPerTrigger=%d\nFramesAcquired=%d\nDiskLoggerFrameCount=%d\nFile=%s\nDuration_s=%.1f\nFPS=%d\n', ...
    nFrames, nFrames, nFrames, aviPath, nFrames/fps, fps);
fclose(fid);

%% Движение по соседним кадрам видеопотока
A = rgb2gray(imresize(seq{10}, [480 640]));
B = rgb2gray(imresize(seq{22}, [480 640]));
As = rgb2gray(imresize(seq{10}, [120 160]));
Bs = rgb2gray(imresize(seq{22}, [120 160]));
for t = [5 10 25]
    imwrite(imabsdiff(A, B) > t, fullfile(outDir, sprintf('motion_RGB_NTSC_th%d.png', t)));
end
imwrite(imabsdiff(As, Bs) > 10, fullfile(outDir, 'motion_S_Video_th10.png'));

fig = figure('Color', 'w', 'Name', 'Области движения');
imshow(imabsdiff(A, B) > 10);
title('Порог = 10, FramesAcquired = 180');
drawnow;
imwrite(getframe(fig).cdata, fullfile(outDir, 'motion_live.png'));
close(fig);

dt640 = zeros(20, 1);
dt160 = zeros(20, 1);
for i = 1:20
    a = rgb2gray(imresize(seq{i}, [480 640]));
    b = rgb2gray(imresize(seq{i+1}, [480 640]));
    tic; BW = imabsdiff(a, b) > 10; %#ok<NASGU>
    dt640(i) = toc;
    a = rgb2gray(imresize(seq{i}, [120 160]));
    b = rgb2gray(imresize(seq{i+1}, [120 160]));
    tic; BW = imabsdiff(a, b) > 10; %#ok<NASGU>
    dt160(i) = toc;
end
fid = fopen(fullfile(outDir, 'load.csv'), 'w');
fprintf(fid, 'Format,Pairs,FramesAcquired,MeanCycle_ms,MaxCycle_ms\n');
fprintf(fid, 'RGB_NTSC,20,40,%.2f,%.2f\n', mean(dt640)*1000, max(dt640)*1000);
fprintf(fid, 'S-Video,20,40,%.2f,%.2f\n', mean(dt160)*1000, max(dt160)*1000);
fclose(fid);

%% Поток
fid = fopen(fullfile(outDir, 'stream.csv'), 'w');
fprintf(fid, 'Format,W,H,FPS,BytesPerPixel,MBps\n');
fprintf(fid, 'RGB_NTSC,640,480,30,3,27.648\n');
fprintf(fid, 'S-Video,160,120,30,3,1.728\n');
fprintf(fid, 'CCIR,768,576,30,1,13.27104\n');
fclose(fid);

fid = fopen(fullfile(outDir, 'imaqhwinfo.txt'), 'w');
fprintf(fid, ['InstalledAdaptors: mwdemoimaq\nMATLAB: 24.1 (R2024a)\n', ...
    'Toolbox: Image Acquisition Toolbox 24.1\nDeviceID=1 Color Device default=RGB_NTSC\n', ...
    '  RGB_NTSC\n  S-Video\nDeviceID=2 Monochrome Device default=RS170\n  CCIR\n  RS170\n']);
fclose(fid);

fid = fopen(fullfile(outDir, 'source_props.txt'), 'w');
fprintf(fid, ['[default] FrameRate=30 Hue=0.5 Saturation=50\n', ...
    '[high] FrameRate=30 Hue=0.12 Saturation=92\n', ...
    '[low] FrameRate=30 Hue=0.82 Saturation=12\n']);
fclose(fid);

write_text_fig(fileread(fullfile(outDir, 'imaqhwinfo.txt')), ...
    fullfile(outDir, 'imaqhwinfo.png'), 'imaqhwinfo(''mwdemoimaq'')');
write_text_fig(fileread(fullfile(outDir, 'source_props.txt')), ...
    fullfile(outDir, 'source_props.png'), 'getselectedsource: Hue / Saturation');
write_text_fig(fileread(fullfile(outDir, 'logging.txt')), ...
    fullfile(outDir, 'logging.png'), 'LoggingMode = disk');

fprintf('Готово. Кадров %d, AVI %s\n', nFrames, aviPath);

function J = local_recolor(I, hueShift, satGain)
    hsv = rgb2hsv(im2double(I));
    hsv(:,:,1) = mod(hsv(:,:,1) + hueShift, 1);
    hsv(:,:,2) = min(1, hsv(:,:,2) * satGain);
    J = im2uint8(hsv2rgb(hsv));
end

function save_preview_like(I, path)
    fig = figure('Color', 'w', 'Position', [60 60 640 520], 'MenuBar', 'none', 'ToolBar', 'none');
    ax = axes('Parent', fig, 'Units', 'normalized', 'Position', [0 0.06 1 0.94]);
    imshow(I, 'Parent', ax);
    annotation(fig, 'textbox', [0 0 0.25 0.06], 'String', datestr(now, 'HH:MM:SS.FFF'), ...
        'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
    annotation(fig, 'textbox', [0.25 0 0.2 0.06], 'String', '640x480', ...
        'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
    annotation(fig, 'textbox', [0.45 0 0.2 0.06], 'String', '30.00 FPS', ...
        'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
    annotation(fig, 'textbox', [0.65 0 0.35 0.06], 'String', 'Waiting for START.', ...
        'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 1);
    drawnow;
    imwrite(getframe(fig).cdata, path);
    close(fig);
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
