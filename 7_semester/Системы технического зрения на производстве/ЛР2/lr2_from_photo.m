%% ЛР2: свой кадр вместо demo-узора
clc; close all;
base = fileparts(mfilename('fullpath'));
outDir = fullfile(base, 'screenshots');
if ~exist(outDir, 'dir'); mkdir(outDir); end
srcPath = fullfile(base, 'moy_kadr.jpg');

I0 = imread(srcPath);
I640 = imresize(I0, [480 640]);
I160 = imresize(I0, [120 160]);
I768 = rgb2gray(imresize(I0, [576 768]));

imwrite(I640, fullfile(outDir, 'snapshot_default.png'));
imwrite(I640, fullfile(outDir, 'mode_RGB_NTSC.png'));
imwrite(I160, fullfile(outDir, 'mode_S_Video.png'));
imwrite(I768, fullfile(outDir, 'mode_CCIR.png'));

fig = figure('Color', 'w', 'Position', [80 80 640 520], 'MenuBar', 'none', 'ToolBar', 'none');
ax = axes('Parent', fig, 'Position', [0 0.06 1 0.94]);
imshow(I640, 'Parent', ax);
annotation(fig, 'textbox', [0 0 0.25 0.06], 'String', datestr(now, 'HH:MM:SS.FFF'), ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 2);
annotation(fig, 'textbox', [0.25 0 0.2 0.06], 'String', '640x480', ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 2);
annotation(fig, 'textbox', [0.45 0 0.2 0.06], 'String', '30.00 FPS', ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 2);
annotation(fig, 'textbox', [0.65 0 0.35 0.06], 'String', 'Waiting for START.', ...
    'EdgeColor', 'k', 'HorizontalAlignment', 'center', 'FontSize', 8, 'Margin', 2);
drawnow;
fr = getframe(fig);
imwrite(fr.cdata, fullfile(outDir, 'preview.png'));
close(fig);

Ihi = local_recolor(I640, -0.08, 1.35);
Ilo = local_recolor(I640, 0.12, 0.45);
imwrite(Ihi, fullfile(outDir, 'snapshot_hue_sat_high.png'));
imwrite(Ilo, fullfile(outDir, 'snapshot_hue_sat_low.png'));
imwrite(Ihi, fullfile(outDir, 'hue_sat_high.png'));
imwrite(Ilo, fullfile(outDir, 'hue_sat_low.png'));

G1 = rgb2gray(I640);
G2 = rgb2gray(imtranslate(I640, [12 4], 'FillValues', 0));
G1s = rgb2gray(I160);
G2s = rgb2gray(imtranslate(I160, [4 2], 'FillValues', 0));

for t = [5 10 25]
    BW = imabsdiff(G1, G2) > t;
    imwrite(BW, fullfile(outDir, sprintf('motion_RGB_NTSC_th%d.png', t)));
end
imwrite(imabsdiff(G1s, G2s) > 10, fullfile(outDir, 'motion_S_Video_th10.png'));

fid = fopen(fullfile(outDir, 'stream.csv'), 'w');
fprintf(fid, 'Format,W,H,FPS,BytesPerPixel,MBps\n');
fprintf(fid, 'RGB_NTSC,640,480,30,3,27.648\n');
fprintf(fid, 'S-Video,160,120,30,3,1.728\n');
fprintf(fid, 'CCIR,768,576,30,1,13.27104\n');
fclose(fid);

fprintf('Кадр обработан: %s\n', srcPath);
fprintf('Crop: %dx%d -> 640x480 / 160x120 / 768x576\n', size(I0,2), size(I0,1));

function J = local_recolor(I, hueShift, satGain)
    hsv = rgb2hsv(im2double(I));
    hsv(:,:,1) = mod(hsv(:,:,1) + hueShift, 1);
    hsv(:,:,2) = min(1, hsv(:,:,2) * satGain);
    J = im2uint8(hsv2rgb(hsv));
end
