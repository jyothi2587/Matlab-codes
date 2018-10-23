clc, clear all; close all;
A = magic(5);
writetable(T,'try1.xls','Sheet',2,'Range','A2:F6')
xlswrite('try1', x, sheet, xlRange);
cHeader = {'PROCHI' 'date' 'hb_extract' 'tablets' 'intakemethod' 'dose' 'hdl' 'chol' 'trigs' 'ldl'}; %header
commaHeader = [cHeader;repmat({','},1,numel(cHeader))]; %insert commaas
commaHeader = commaHeader(:)';
textHeader = cell2mat(commaHeader); %cHeader in text with commas
%write header to file
fid = fopen('try1.xls','w'); 
fprintf(fid,'%s\n',textHeader);
fclose(fid);
% generate random gaussian numbers for hdl
filename = 'testdata.xlsx';
xlswrite(testdata,A,4,'D1:D411');
rng(0,'twister');
a = 0.4; % std dev
b = 1.3; % mean
y = a.*randn(1000,411) + b;
stats = [mean(y) std(y) var(y)];
% generate random gaussian numbers for chol
filename = 'testdata.xlsx';
xlswrite(testdata,A,5,'E1:E411');
rng(0,'twister');
a = 1.2; % std dev
b = 4.6; % mean 
y = a.*randn(1000,411) + b;
stats = [mean(y) std(y) var(y)];
% generate random gaussian numbers for trigs
filename = 'testdata.xlsx';
xlswrite(testdata,A,6,'F1:F411');
rng(0,'twister');
a = 1.7; % std dev
b = 2.9; % mean
y = a.*randn(1000,411) + b;
stats = [mean(y) std(y) var(y)];
% generate random gaussian numbers for ldl
filename = 'testdata.xlsx';
xlswrite(testdata,A,7,'G1:G411');
rng(0,'twister');
a = 1; % std dev
b = 2.4; % mean
y = a.*randn(1000,411) + b;
stats = [mean(y) std(y) var(y)];
% generate random gaussian numbers for hba1c_ifcc
filename = 'testdata.xlsx';
xlswrite(testdata,A,8,'H1:H411');
rng(0,'twister');
a = 17.63; % std dev
b = 59.69; % mean
y = a.*randn(1000,411) + b;
stats = [mean(y) std(y) var(y)];
% generate random gaussian numbers for hba1c_dcct
rng(0,'twister');
a = 1.61; % std dev
b = 7.61; % mean
y = a.*randn(1000,411) + b;
stats = [mean(y) std(y) var(y)];
filename = 'testdata.xlsx';
xlswrite(testdata,A,9,'I1:I411');
