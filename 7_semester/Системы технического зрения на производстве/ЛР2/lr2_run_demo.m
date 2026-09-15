%% ЛР2 на демонстрационном адаптере MathWorks (mwdemoimaq)
% Физической web-камеры на ПК нет; используется штатный demo-адаптор IAT.

clc; close all; imaqreset;
base = fileparts(mfilename('fullpath'));
outDir = fullfile(base, 'screenshots');
if ~exist(outDir, 'dir'); mkdir(outDir); end

dll = fullfile(matlabroot, 'toolbox', 'imaq', 'imaqadaptors', 'kit', 'demo', 'win64', 'mwdemoimaq.dll');
regs = imaqregister;
if ~any(strcmp(regs, dll))
    imaqregister(dll);
end
imaqreset;

info = imaqhwinfo('mwdemoimaq');
fid = fopen(fullfile(outDir, 'camera_info.txt'), 'w');
fprintf(fid, 'Adaptor: mwdemoimaq\nMATLAB: %s\nDevices: %d\n', version, numel(info.DeviceIDs));
for i = 1:numel(info.DeviceIDs)
    cam = imaqhwinfo('mwdemoimaq', info.DeviceIDs{i});
    fprintf(fid, '\nID=%d  %s  default=%s\n', cam.DeviceID, cam.DeviceName, cam.DefaultFormat);
    fprintf(fid, '  %s\n', cam.SupportedFormats{:});
end
fclose(fid);

cfgs = {
    1, 'RGB_NTSC'
    1, 'S-Video'
    2, 'CCIR'
};
Bmap = [3; 3; 1]; % RGB, RGB, grayscale

%% Preview + свойства
vid = videoinput('mwdemoimaq', 1, 'RGB_NTSC');
vid.ReturnedColorSpace = 'rgb';
preview(vid);
pause(2);
fr = getframe(gcf);
imwrite(fr.cdata, fullfile(outDir, 'preview.png'));
I0 = getsnapshot(vid);
imwrite(I0, fullfile(outDir, 'snapshot_default.png'));

src = getselectedsource(vid);
props = get(src);
writecell(fieldnames(props), fullfile(outDir, 'source_prop_names.txt'));

stoppreview(vid); closepreview(vid);
src.Hue = 0.15;
src.Saturation = 90;
preview(vid); pause(2);
fr = getframe(gcf);
imwrite(fr.cdata, fullfile(outDir, 'hue_sat_high.png'));
I1 = getsnapshot(vid);
imwrite(I1, fullfile(outDir, 'snapshot_hue_sat_high.png'));

stoppreview(vid); closepreview(vid);
src.Hue = 0.85;
src.Saturation = 10;
preview(vid); pause(2);
fr = getframe(gcf);
imwrite(fr.cdata, fullfile(outDir, 'hue_sat_low.png'));
I2 = getsnapshot(vid);
imwrite(I2, fullfile(outDir, 'snapshot_hue_sat_low.png'));
stoppreview(vid); closepreview(vid);
delete(vid); imaqreset;

%% Поток данных
rows = {};
for i = 1:size(cfgs, 1)
    vid = videoinput('mwdemoimaq', cfgs{i, 1}, cfgs{i, 2});
    src = getselectedsource(vid);
    res = vid.VideoResolution;
    fps = str2double(string(src.FrameRate));
    if isnan(fps); fps = 30; end
    B = Bmap(i);
    D = res(1) * res(2) * B * fps / 1e6;
    frame = getsnapshot(vid);
    imwrite(frame, fullfile(outDir, sprintf('mode_%s.png', regexprep(cfgs{i, 2}, '\W', '_'))));
    rows(end+1, :) = {cfgs{i, 2}, res(1), res(2), fps, B, D}; %#ok<SAGROW>
    delete(vid); imaqreset;
end
T = cell2table(rows, 'VariableNames', {'Format', 'W', 'H', 'FPS', 'BytesPerPixel', 'MBps'});
writetable(T, fullfile(outDir, 'stream.csv'));
disp(T);

%% Детект движения на двух разрешениях
motion_run(1, 'RGB_NTSC', [5 10 25], outDir, 40);
motion_run(1, 'S-Video', [10], outDir, 40);

imaqreset; close all;
fprintf('Saved to %s\n', outDir);

function motion_run(devId, fmt, thresholds, outDir, nNeed)
    for t = thresholds
        vid = videoinput('mwdemoimaq', devId, fmt);
        vid.ReturnedColorSpace = 'rgb';
        vid.TriggerRepeat = Inf;
        vid.FramesPerTrigger = 20;
        start(vid);
        last = [];
        saved = false;
        n = 0;
        tic;
        while isrunning(vid) && toc < 6
            if vid.FramesAvailable >= 1
                data = getdata(vid, 1);
                I = data(:,:,:,1);
                if size(I, 3) == 3
                    I = rgb2gray(I);
                end
                if ~isempty(last)
                    BW = imabsdiff(I, last) > t;
                    n = n + 1;
                    if ~saved
                        imwrite(BW, fullfile(outDir, sprintf('motion_%s_th%d.png', regexprep(fmt, '\W', '_'), t)));
                        saved = true;
                    end
                end
                last = I;
            end
            if n >= nNeed
                break;
            end
        end
        stop(vid); delete(vid); imaqreset;
    end
end
