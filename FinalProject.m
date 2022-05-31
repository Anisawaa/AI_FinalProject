function varargout = FinalProject(varargin)
% FINALPROJECT MATLAB code for FinalProject.fig
%      FINALPROJECT, by itself, creates a new FINALPROJECT or raises the existing
%      singleton*.
%
%      H = FINALPROJECT returns the handle to a new FINALPROJECT or the handle to
%      the existing singleton*.
%
%      FINALPROJECT('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in FINALPROJECT.M with the given input arguments.
%
%      FINALPROJECT('Property','Value',...) creates a new FINALPROJECT or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before FinalProject_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to FinalProject_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help FinalProject

% Last Modified by GUIDE v2.5 26-May-2022 21:55:31

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @FinalProject_OpeningFcn, ...
                   'gui_OutputFcn',  @FinalProject_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before FinalProject is made visible.
function FinalProject_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to FinalProject (see VARARGIN)

% Choose default command line output for FinalProject
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes FinalProject wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = FinalProject_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in btntable.
function btntable_Callback(hObject, eventdata, handles)
% hObject    handle to btntable (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
dataset = readcell('cosmetics.csv', 'Range', 'B2:K299');
header = readcell('cosmetics.csv', 'Range', 'B1:K1');

set (handles.table, 'Data', dataset, 'ColumnName', header);

% --- Executes on selection change in popmenu.
function popmenu_Callback(hObject, eventdata, handles)
% hObject    handle to popmenu (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in btnresult.
function btnresult_Callback(hObject, eventdata, handles)
% hObject    handle to btnresult (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
p_r = str2double(get(handles.imp1, 'string'));
r_s = str2double(get(handles.imp2, 'string'));
s_p = str2double(get(handles.imp3, 'string'));

contents = cellstr(get(handles.popmenu, 'string'))
skintype = contents(get(handles.popmenu, 'Value'))

%% Calculation of Paired Comparison Matrix on Criteria
    
    %        P       R       S    
    PCM = [  1      p_r    1/s_p; %P - Price
           1/p_r     1      r_s;  %R - Rank
            s_p    1/r_s     1 ]  %S - Skin Type
        
    %Normalization
    disp('Normalization')
    wPCM = calc_norm(PCM)
    
    %Calculation of Eigen Vector
    [m n] = size(wPCM);
    for i = 1 : m,  
        sumRow = 0;  
        for j = 1 : n,  
            sumRow = sumRow + wPCM(i, j);  
        end;
        
    V(i) = (sumRow);
    end;
    
    disp('Eigen Criteria')
    wPCM = transpose(V) / m
        
%% Calculation of Price Alternative
    
    opts = detectImportOptions('cosmetics.csv');
    opts.SelectedVariableNames = (4);
    PQAp = readmatrix('cosmetics.csv', opts);  
    
    %Normalization
    wPrice = calc_norm(PQAp);

 %% Calculation of Rank Alternative
    
    opts = detectImportOptions('cosmetics.csv');
    opts.SelectedVariableNames = (5);
    PQAr = readmatrix('cosmetics.csv', opts);   
    
    %Normalization
    wRank = calc_norm(PQAr); 
 
%% Calculation of Skin Type Alternative
    
    if (strcmp(skintype, 'Combination'))
        opts = detectImportOptions('cosmetics.csv')
        opts.SelectedVariableNames = (7)
        PQAs = readmatrix('cosmetics.csv', opts)

    elseif (strcmp(skintype, 'Dry'))
        opts = detectImportOptions('cosmetics.csv')
        opts.SelectedVariableNames = (8)
        PQAs = readmatrix('cosmetics.csv', opts)

    elseif (strcmp(skintype, 'Normal'))
        opts = detectImportOptions('cosmetics.csv')
        opts.SelectedVariableNames = (9)
        PQAs = readmatrix('cosmetics.csv', opts)

    elseif (strcmp(skintype, 'Oily'))
        opts = detectImportOptions('cosmetics.csv')
        opts.SelectedVariableNames = (10)
        PQAs = readmatrix('cosmetics.csv', opts)

    elseif (strcmp(skintype, 'Sensitive'))
        opts = detectImportOptions('cosmetics.csv')
        opts.SelectedVariableNames = (11)
        PQAs = readmatrix('cosmetics.csv', opts)
    end
    
    %Normalization
    wSkin = calc_norm(PQAs);
    
%% Determine the Final Result
    wM = [wPrice wRank wSkin];

    %Multiply PCM and each PQA to get the Final Score
    MC_Scores = wM * wPCM;

    %Showcase the Final Result
    opts = detectImportOptions('cosmetics.csv');
    opts.SelectedVariableNames = (3);
    brand = readmatrix('cosmetics.csv', opts);
    
    mc_max = max(MC_Scores)
    rowValue = find(MC_Scores == mc_max)
    resultValue = brand(rowValue)
    
    set(handles.row, 'String', rowValue);
    set(handles.result, 'String', resultValue);

    function [normvect ] = calc_norm(M)
        sM = sum(M);
        normvect = M./sM;
        
% --- Executes on button press in btnreset.
function btnreset_Callback(hObject, eventdata, handles)
% hObject    handle to btnreset (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
set(handles.table, 'Data', '');
set(handles.imp1, 'string', '');
set(handles.imp2, 'string', '');
set(handles.imp3, 'string', '');
set(handles.popmenu, 'Value', 1);
set(handles.result, 'string', '');
set(handles.row, 'string', '');
clc;

function imp2_Callback(hObject, eventdata, handles)
% hObject    handle to imp2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of imp2 as text
%        str2double(get(hObject,'String')) returns contents of imp2 as a double


% --- Executes during object creation, after setting all properties.
function imp2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to imp2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function imp1_Callback(hObject, eventdata, handles)
% hObject    handle to imp1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of imp1 as text
%        str2double(get(hObject,'String')) returns contents of imp1 as a double


% --- Executes during object creation, after setting all properties.
function imp1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to imp1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end




% --- Executes during object creation, after setting all properties.
function popmenu_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popmenu (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function imp3_Callback(hObject, eventdata, handles)
% hObject    handle to imp3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of imp3 as text
%        str2double(get(hObject,'String')) returns contents of imp3 as a double


% --- Executes during object creation, after setting all properties.
function imp3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to imp3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function row_CreateFcn(hObject, eventdata, handles)
% hObject    handle to row (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
