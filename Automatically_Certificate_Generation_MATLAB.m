
% Clear the command window, workspace, and close all figures
clc 
clear all 
close all 
home 

% Read data from an Excel file called 'Registration_Details.xls' and store it in two arrays
filename = 'Registration_Details.xls';
[num,txt] = xlsread(filename)

% Determine the length of the 'txt' array
len=length(txt)

% Load a blank certificate image
blankimage = imread('Certificate_Sample.tif');

% Extract the name data from 'txt' and store it in a new array called 'text_names'
for i=1:len
    for j= 2:2 
        text_names(i,j)=txt(i,j)
    end
end

% Extract the topic data from 'txt' and store it in a new array called 'text_topic'
for i=1:len
    for j= 3:3
      text_topic(i,j)=txt(i,j)
    end
end

% Generate a certificate for each person in the 'txt' array
for i=2:len % Start the loop from 2 because the first row of 'txt' contains column headers
        % Combine the name and topic data for each person into a single string
        text_all=[text_names(i,2) text_topic(i,3)]
        % Define the position where the text will be inserted into the certificate image
        position = [331 334;308 430];          
        % Insert the text into the certificate image
        RGB = insertText(blankimage,position,text_all,'FontSize',22,'BoxOpacity',0);        
        % Display the certificate image
        figure
        imshow(RGB)               
        % Save the certificate image to a file
        y=i-1
        filename=['CertificateNo_' num2str(y)]
        lastf=[filename '.tif']
        saveas(gcf,lastf)
end 
