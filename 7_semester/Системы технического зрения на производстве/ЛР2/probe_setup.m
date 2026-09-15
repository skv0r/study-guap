%% ЛР2: проверка MATLAB, адаптера winvideo и камеры
clc; close all;

fprintf('MATLAB %s\n', version);
v = ver;
need = {'Image Acquisition Toolbox', 'Image Processing Toolbox'};
names = {v.Name};
for k = 1:numel(need)
    if any(strcmp(names, need{k}))
        fprintf('OK  %s\n', need{k});
    else
        fprintf('NO  %s  — поставьте toolbox\n', need{k});
    end
end

imaqreset;
info = imaqhwinfo;
disp('InstalledAdaptors:');
disp(info.InstalledAdaptors);

if isempty(info.InstalledAdaptors)
    error(['Нет адаптера камеры. В MATLAB: Home → Add-Ons → Get Add-Ons, ', ...
           'установите Image Acquisition Toolbox Support Package for OS Generic Video Interface, ', ...
           'перезапустите MATLAB.']);
end

if ~any(strcmp(info.InstalledAdaptors, 'winvideo'))
    error('Адаптер winvideo не найден. Нужен Support Package OS Generic Video Interface.');
end

w = imaqhwinfo('winvideo');
disp(w);
if isempty(w.DeviceIDs)
    error(['Камера не видна MATLAB. Закройте Zoom/Teams/Discord/браузер, ', ...
           'подключите USB-вебку или штатную камеру ноутбука, затем imaqreset.']);
end

for i = 1:numel(w.DeviceIDs)
    id = w.DeviceIDs{i};
    cam = imaqhwinfo('winvideo', id);
    fprintf('DeviceID=%d  %s  default=%s  formats=%d\n', ...
        id, cam.DeviceName, cam.DefaultFormat, numel(cam.SupportedFormats));
end

fprintf('Проверка среды пройдена.\n');
