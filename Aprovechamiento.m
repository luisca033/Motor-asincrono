function varargout = Aprovechamiento(varargin)

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Aprovechamiento_OpeningFcn, ...
                   'gui_OutputFcn',  @Aprovechamiento_OutputFcn, ...
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


% --- Executes just before Aprovechamiento is made visible.
function Aprovechamiento_OpeningFcn(hObject, eventdata, handles, varargin)
function log_CreateFcn(hObject, eventdata, handles)
log=imread('salle.jpg');
imshow(log);

% Hint: place code in OpeningFcn to populate log


% --- Executes during object creation, after setting all properties.
function apro_CreateFcn(hObject, eventdata, handles)
ap=imread('apro.jpg');
imshow(ap);
% Hint: place code in OpeningFcn to populate apro


% --- Executes during object creation, after setting all properties.
function mor_CreateFcn(hObject, eventdata, handles)
mor=imread('motor.jpg');
imshow(mor);
% Hint: place code in OpeningFcn to populate mor


% --- Executes during object creation, after setting all properties.
function T_CreateFcn(hObject, eventdata, handles)
t=imread('t.jpg');
imshow(t);
% Hint: place code in OpeningFcn to populate T


% --- Executes during object creation, after setting all properties.
function L_CreateFcn(hObject, eventdata, handles)
l=imread('l.jpg');
imshow(l);


function sin_p_CreateFcn(hObject, eventdata, handles)
tsin=imread('t_sin.jpg');
imshow(tsin);

% Choose default command line output for Aprovechamiento
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Aprovechamiento wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Aprovechamiento_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes during object creation, after setting all properties.


% --------------------------------------------------------------------
function Datos_Callback(hObject, eventdata, handles)
% hObject    handle to Datos (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Abrir_datos_Callback(hObject, eventdata, handles)
% hObject    handle to Abrir_datos (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
winopen('parametros.xlsx')


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% Exportar datos

num=num2str(xlsread('parametros.xlsx','Datos_ope','B21'))
num2=num
% Datos Operacion
Pnom=xlsread('parametros.xlsx','Datos_ope','B3')
fnom=xlsread('parametros.xlsx','Datos_ope','B5')
pol=xlsread('parametros.xlsx','Datos_ope','B6')
unom=xlsread('parametros.xlsx','Datos_ope','B7')
inom=xlsread('parametros.xlsx','Datos_ope','B8')
fpdop=xlsread('parametros.xlsx','Datos_ope','B9')
Hop=xlsread('parametros.xlsx','Datos_ope','B10')
ks=xlsread('parametros.xlsx','Datos_ope','B11')
kr=xlsread('parametros.xlsx','Datos_ope','B12')
gm=xlsread('parametros.xlsx','Datos_ope','B13')
r11=xlsread('parametros.xlsx','Datos_ope','B14')
O1=xlsread('parametros.xlsx','Datos_ope','B15')
Onl=xlsread('parametros.xlsx','Datos_ope','B16')
raiz=sqrt(3)
pfw=xlsread('parametros.xlsx','Datos_ope','B17')
ko=xlsread('parametros.xlsx','Datos_ope','B18')
ki=xlsread('parametros.xlsx','Datos_ope','B19')
ufin=xlsread('parametros.xlsx','Datos_ope','J3')
ifin=xlsread('parametros.xlsx','Datos_ope','J4')
pfin=xlsread('parametros.xlsx','Datos_ope','J5')
O2=xlsread('parametros.xlsx','Datos_ope','J6')
n=xlsread('parametros.xlsx','Datos_ope','J7')
nsyn=xlsread('parametros.xlsx','Datos_ope','J8')
cosy=xlsread('parametros.xlsx','Datos_ope','J9')
lry=xlsread('parametros.xlsx','Datos_ope','J10')
% Resistencias perdida estator
u1=xlsread('parametros.xlsx','Resistencias perdida estator',strcat('A3:A',num,''))
i1=xlsread('parametros.xlsx','Resistencias perdida estator',strcat('B3:B',num,''))
p1=xlsread('parametros.xlsx','Resistencias perdida estator',strcat('C3:C',num,''))

% Inductancia en el estator
u2=xlsread('parametros.xlsx','Inductancia del estator',strcat('A3:A',num,''))
i2=xlsread('parametros.xlsx','Inductancia del estator',strcat('B3:B',num,''))
p2=xlsread('parametros.xlsx','Inductancia del estator',strcat('C3:C',num,''))

% Inductancia de dispercion
u3=xlsread('parametros.xlsx','Inductancia de dispercion',strcat('A3:A',num,''))
i3=xlsread('parametros.xlsx','Inductancia de dispercion',strcat('B3:B',num,''))
p3=xlsread('parametros.xlsx','Inductancia de dispercion',strcat('C3:C',num,''))

% Inductancia magnetizante
u4=xlsread('parametros.xlsx','Inductancia de dispercion',strcat('A3:A',num,''))
i4=xlsread('parametros.xlsx','Inductancia de dispercion',strcat('B3:B',num,''))
p4=xlsread('parametros.xlsx','Inductancia de dispercion',strcat('C3:C',num,''))

seleccion=handles.edit1;
if seleccion==1
    % prueba en y
% Estrella
% Resistencia del devanado del estator
numrs1=ks+25;
denrs1=ks+O1;
rs1=numrs1/denrs1;
rs=(1/2)*r11*rs1;

% Perdidas entre hierro
is=i1
is2=i1.*i1;
pk=p1-(3.*is2.*rs.*((ks+Onl)/(ks+25)))
up2=u1.*u1
pfe=pk-pfw
%Calculo de Rfe
rfe= 3.*(up2./pfe)
iim1=xlswrite('parametros.xlsx',is,'Resistencias perdida estator',strcat('D3:D',num,''))
pim1=xlswrite('parametros.xlsx',pk,'Resistencias perdida estator',strcat('E3:E',num,''))
uim1=xlswrite('parametros.xlsx',up2,'Resistencias perdida estator',strcat('F3:F',num,''))
pim12=xlswrite('parametros.xlsx',pfe,'Resistencias perdida estator',strcat('G3:G',num,''))
rfeim=xlswrite('parametros.xlsx',rfe,'Resistencias perdida estator',strcat('H3:H',num,''))
%
% Inductancia total del estator
% Impedancia total
zs=(u2./(raiz.*i2))
fp=(p2./(u2.*i2.*raiz))
r2=zs.*fp
im=i2
xn=sqrt((zs.*zs)-(r2.*r2))
lts=xn./(2*pi*fnom)


zsm=xlswrite('parametros.xlsx',zs,'Inductancia del estator',strcat('D3:D',num,''))
fpm=xlswrite('parametros.xlsx',fp,'Inductancia del estator',strcat('E3:E',num,''))
r2m=xlswrite('parametros.xlsx',r2,'Inductancia del estator',strcat('F3:F',num,''))
iim2=xlswrite('parametros.xlsx',im,'Inductancia del estator',strcat('G3:G',num,''))
xm2=xlswrite('parametros.xlsx',xn,'Inductancia del estator',strcat('H3:H',num,''))
ltsm=xlswrite('parametros.xlsx',lts,'Inductancia del estator',strcat('I3:I',num,''))

% Inductancia de dispercion

zs2=(u3./(raiz.*i3))
fp1=(p3./(u3.*i3.*raiz))
r3=zs2.*fp1
im1=i3
xc=sqrt((zs2.*zs2)-(r3.*r3))
lcs=xc./(2*pi*fnom)


zsm2=xlswrite('parametros.xlsx',zs2,'Inductancia de dispercion',strcat('D3:D',num,''))
fpm1=xlswrite('parametros.xlsx',fp1,'Inductancia de dispercion',strcat('E3:E',num,''))
r3m=xlswrite('parametros.xlsx',r3,'Inductancia de dispercion',strcat('F3:F',num,''))
iim3=xlswrite('parametros.xlsx',im1,'Inductancia de dispercion',strcat('G3:G',num,''))
xm3=xlswrite('parametros.xlsx',xc,'Inductancia de dispercion',strcat('H3:H',num,''))
lcsm=xlswrite('parametros.xlsx',lcs,'Inductancia de dispercion',strcat('I3:I',num,''))

% Inductancia magnetizante

h=(Hop/1000)*(0.21-(pol/100));
%gamma
gam=h*sqrt(pi*2*fnom*4*pi*10^-7);
%ki=((3/2)*gam)*(((sinh(2*gam)-sin(2*gam))/((cosh(2*gam)-cos(2*gam)))));
%ko=lts./lcs; % preguntar
num=ko+1;
den=ko+ki;
ah=num./den;
lo=lcs.*ah
inv=1./ko;
lm=lts-(lo./(1+inv))
um=2*pi*fnom.*lm.*im
lsout=lts-lm
lrout=lo-lsout

im4=xlswrite('parametros.xlsx',i2,'Inductancia magnetizante',strcat('A3:A',num2));
ltsm2n=xlswrite('parametros.xlsx',lts,'Inductancia magnetizante',strcat('B3:B',num2))
lomn=xlswrite('parametros.xlsx',lo,'Inductancia magnetizante',strcat('C3:C',num2))
lmmn=xlswrite('parametros.xlsx',lm,'Inductancia magnetizante',strcat('D3:D',num2))
lsoutmn=xlswrite('parametros.xlsx',lsout,'Inductancia magnetizante',strcat('E3:E',num2))
lroutn=xlswrite('parametros.xlsx',lrout,'Inductancia magnetizante',strcat('F3:F',num2))
ummn=xlswrite('parametros.xlsx',um,'Inductancia magnetizante',strcat('G3:G',num2))
%
% Flujo constante
ifi=ifin;
uf=ufin/raiz;
lsprom=mean(lsout);
he=(ks+O2)/(ks+25);
uma=uf-(ifi*(cosy*rs+sqrt(1+(cosy*cosy))*2*pi*lsprom*fnom));% norma
umb=ifi*(sqrt(1-(cosy*cosy))*rs-(cosy*2*pi*lsprom*fnom));% norma
umf=sqrt((uma*uma)+(umb*umb));
lmprom=mean(lm);
iout=((umb/(2*pi*fnom*lmprom)-ifi*cosy)^2+(ifi*sqrt(1-(cosy*cosy)-(uma/(2*pi*fnom*lmprom)))^2)); % norma
s=(nsyn-n)/nsyn;
zf=ufin/(ifi*raiz);
xrprim=2*pi*fnom*lry;
xsprim=2*pi*fnom*lsprom;
xmprim=2*pi*fnom*lmprom;
xout=zf*sqrt(1-(cosy*cosy));
rpim=s*(xrprim+xmprim)*sqrt(((xrprim*xmprim)-(xout-xsprim)/(xrprim+xmprim))/(xout-xsprim-xmprim))*((kr+25)/(kr+O2));
a=xlswrite('parametros.xlsx',uf,'Con Carga','B2')
b=xlswrite('parametros.xlsx',ifi,'Con Carga','B3')
c=xlswrite('parametros.xlsx',lsprom,'Con Carga','B4')
d=xlswrite('parametros.xlsx',cosy,'Con Carga','B5')
e=xlswrite('parametros.xlsx',nsyn,'Con Carga','B6')
f=xlswrite('parametros.xlsx',s,'Con Carga','B7')
g=xlswrite('parametros.xlsx',uma,'Con Carga','B8')
j=xlswrite('parametros.xlsx',umb,'Con Carga','B9')
k=xlswrite('parametros.xlsx',umf,'Con Carga','B10')
l=xlswrite('parametros.xlsx',lmprom,'Con Carga','B11')
m=xlswrite('parametros.xlsx',xrprim,'Con Carga','B12')
nm=xlswrite('parametros.xlsx',xsprim,'Con Carga','B13')
op=xlswrite('parametros.xlsx',xmprim,'Con Carga','B14')
p=xlswrite('parametros.xlsx',iout,'Con Carga','B15')
q=xlswrite('parametros.xlsx',lry,'Con Carga','B16')
r=xlswrite('parametros.xlsx',zf,'Con Carga','B17')
st=xlswrite('parametros.xlsx',xout,'Con Carga','B18')
%t=xlswrite('parametros.xlsx',rpim,'Con Carga','B19');
    
    d = dialog('Position',[300 300 250 150],'Name','Dialogo');

    txt = uicontrol('Parent',d,...
               'Style','text',...
               'Position',[20 80 210 40],...
               'String','programa ejecutado puede abrir el archivo de excel');

    btn = uicontrol('Parent',d,...
               'Position',[85 20 70 25],...
               'String','Close',...
               'Callback','delete(gcf)');

elseif seleccion==0
    % prueba en delta
    % Delta
% Resistencia del devanado del estator
rs1=((ks+25)/(ks+O1));
rs=(3/2)*r11*rs1;
% Perdidas entre hierro
is=i1./raiz;
is2=i1.*i1;
pk=p1-(3.*is2.*rs.*((ks+Onl)/(ks+25)));
up2=u1.*u1;
pfe=pk-pfw;
iim1=xlswrite('parametros.xlsx',is,'Resistencias perdida estator',strcat('D3:D',num,''));
pim1=xlswrite('parametros.xlsx',pk,'Resistencias perdida estator',strcat('E3:E',num,''));
uim1=xlswrite('parametros.xlsx',up2,'Resistencias perdida estator',strcat('F3:F',num,''));
pim12=xlswrite('parametros.xlsx',pfe,'Resistencias perdida estator',strcat('g3:g',num,''));

%Calculo de Rfe
us=u2./raiz;
rfe= 3.*(u2./pfe);
rfeim=xlswrite('parametros.xlsx',rfe,'Resistencias perdida estator',strcat('H3:H',num,''));
% Inductancia total del estator
% Impedancia total
zs=((raiz.*u2)./i2);
fp=(p2./(u2.*i2.*raiz));
r2=zs.*fp;
im=i2./raiz;
xn=sqrt((zs.*zs)-(r2.*r2));
lts=xn./(2*pi*fnom);

zsm=xlswrite('parametros.xlsx',zs,'Inductancia del estator',strcat('D3:D',num,''));
fpm=xlswrite('parametros.xlsx',fp,'Inductancia del estator',strcat('E3:E',num,''));
r2m=xlswrite('parametros.xlsx',r2,'Inductancia del estator',strcat('F3:F',num,''));
iim2=xlswrite('parametros.xlsx',im,'Inductancia del estator',strcat('G3:G',num,''));
xm2=xlswrite('parametros.xlsx',xn,'Inductancia del estator',strcat('H3:H',num,''));
ltsm=xlswrite('parametros.xlsx',lts,'Inductancia del estator',strcat('I3:I',num,''));


% Inductancia de dispercion

zs2=((raiz.*u3)./i3);
fp1=(p3./(u3.*i3.*raiz));
r3=zs2.*fp1;
im1=i3./raiz;
xc=sqrt((zs2.*zs2)-(r3.*r3));
lcs=xc./(2*pi*fnom);


zsm2=xlswrite('parametros.xlsx',zs2,'Inductancia de dispercion',strcat('D3:D',num,''));
fpm1=xlswrite('parametros.xlsx',fp1,'Inductancia de dispercion',strcat('E3:E',num,''));
r3m=xlswrite('parametros.xlsx',r3,'Inductancia de dispercion',strcat('F3:F',num,''));
iim3=xlswrite('parametros.xlsx',im1,'Inductancia de dispercion',strcat('G3:G',num,''));
xm3=xlswrite('parametros.xlsx',xc,'Inductancia de dispercion',strcat('H3:H',num,''));
lcsm=xlswrite('parametros.xlsx',lcs,'Inductancia de dispercion',strcat('I3:I',num,''));

% Inductancia magnetizante

h=(Hop/1000)*(0.21-pol/100);
%gamma
gam=h*sqrt(pi*2*fnom*4*pi*10^-7);
%ki=(3/2*gam)*((sinh(2*gam)-sin(2*gam)/(cosh(2*gam)-cos(2*gam))));
%ko=lts./lcs; % preguntar
num=ko+1;
den=ko+ki;
ah=num./den;
lo=lcs.*ah;
inv=1./ko;
lm=lts-(lo./(1+inv));
um=2*pi*fnom.*lm.*im;
lsout=lts-lm;
lrout=lo-lsout;

im4=xlswrite('parametros.xlsx',im,'Inductancia magnetizante',strcat('A3:A',num,''));
ltsm2=xlswrite('parametros.xlsx',lts,'Inductancia magnetizante',strcat('B3:B',num,''));
lom=xlswrite('parametros.xlsx',lo,'Inductancia magnetizante',strcat('C3:C',num,''));
lmm=xlswrite('parametros.xlsx',lm,'Inductancia magnetizante',strcat('D3:D',num,''));
lsoutm=xlswrite('parametros.xlsx',lsout,'Inductancia magnetizante',strcat('E3:E',num,''));
lrout=xlswrite('parametros.xlsx',lrout,'Inductancia magnetizante',strcat('F3:F',num,''));
umm=xlswrite('parametros.xlsx',um,'Inductancia magnetizante',strcat('G3:G',num,''));

% Flujo constante
ifi=ifin/raiz;
uf=ufin;

lsprom=mean(lsout);
he=(ks+O2)/(ks+25);
uma=uf-(ifi*(cosy*rs+sqrt(1+(cosy*cosy))*2*pi*lsprom*fnom));% norma
umb=ifi*(sqrt(1-(cosy*cosy))*rs-(cosy*2*pi*lsprom*fnom));% norma
umf=sqrt((uma*uma)+(umb*umb));
lmprom=mean(lm);
iout=((umb/(2*pi*fnom*lmprom)-ifi*cosy)^2+(ifi*sqrt(1-(cosy*cosy)-(uma/(2*pi*fnom*lmprom)))^2)); % norma
s=(nsyn-n)/nsyn;
zf=ufin/(ifi*raiz);
xrprim=2*pi*fnom*lry;
xsprim=2*pi*fnom*lsprom;
xmprim=2*pi*fnom*lmprom;
xout=zf*sqrt(1-(cosy*cosy));
rpim=s*(xrprim+xmprim)*sqrt(((xrprim*xmprim)-(xout-xsprim)/(xrprim+xmprim))/(xout-xsprim-xmprim))*((kr+25)/(kr+O2));
a=xlswrite('parametros.xlsx',uf,'Con Carga','B2');
b=xlswrite('parametros.xlsx',ifi,'Con Carga','B3');
c=xlswrite('parametros.xlsx',lsprom,'Con Carga','B4');
d=xlswrite('parametros.xlsx',cosy,'Con Carga','B5');
e=xlswrite('parametros.xlsx',nsyn,'Con Carga','B6');
f=xlswrite('parametros.xlsx',s,'Con Carga','B7');
g=xlswrite('parametros.xlsx',uma,'Con Carga','B8');
j=xlswrite('parametros.xlsx',umb,'Con Carga','B9');
k=xlswrite('parametros.xlsx',umf,'Con Carga','B10');
l=xlswrite('parametros.xlsx',lmprom,'Con Carga','B11');
m=xlswrite('parametros.xlsx',xrprim,'Con Carga','B12');
nm=xlswrite('parametros.xlsx',xsprim,'Con Carga','B13');
op=xlswrite('parametros.xlsx',xmprim,'Con Carga','B14');
p=xlswrite('parametros.xlsx',iout,'Con Carga','B15');
q=xlswrite('parametros.xlsx',lry,'Con Carga','B16');
r=xlswrite('parametros.xlsx',zf,'Con Carga','B17');
st=xlswrite('parametros.xlsx',xout,'Con Carga','B18');
%t=xlswrite('parametros.xlsx',rpim,'Con Carga','B19');

    d = dialog('Position',[300 300 250 150],'Name','Dialogo');

    txt = uicontrol('Parent',d,...
               'Style','text',...
               'Position',[20 80 210 40],...
               'String','programa ejecutado puede abrir el archivo de excel');

    btn = uicontrol('Parent',d,...
               'Position',[85 20 70 25],...
               'String','Close',...
               'Callback','delete(gcf)');
    
else
    display('entrada no identificada no se hara nada')
     d = dialog('Position',[300 300 250 150],'Name','Dialogo');

    txt = uicontrol('Parent',d,...
               'Style','text',...
               'Position',[20 80 210 40],...
               'String','Ingrese 1 o 0 para ejecutar el programa');

    btn = uicontrol('Parent',d,...
               'Position',[85 20 70 25],...
               'String','Close',...
               'Callback','delete(gcf)');
end



function edit1_Callback(hObject, eventdata, handles)

val=get(hObject,'String');
realval=str2double(val);
handles.edit1=realval;
guidata(hObject,handles);

% --- Executes during object creation, after setting all properties.



function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
val1=get(hObject,'String');
realval1=str2double(val1);
handles.edit2=realval1;
guidata(hObject,handles);
% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function axes8_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes8 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
back=imread('salle.JPG');
imshow(back);
% Hint: place code in OpeningFcn to populate axes8


% --- Executes during object creation, after setting all properties.
function axes10_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
back=imread('t_sin.JPG');
imshow(back);
% Hint: place code in OpeningFcn to populate axes10


% --- Executes during object creation, after setting all properties.
function axes11_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes11 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
back=imread('t.JPG');
imshow(back);
% Hint: place code in OpeningFcn to populate axes11


% --- Executes during object creation, after setting all properties.
function axes12_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
back=imread('l.JPG');
imshow(back);
% Hint: place code in OpeningFcn to populate axes12


% --- Executes during object creation, after setting all properties.
function axes9_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes9 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
back=imread('mode.JPG');
imshow(back);
% Hint: place code in OpeningFcn to populate axes9
