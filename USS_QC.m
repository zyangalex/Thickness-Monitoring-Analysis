function [PF,Action] = USS_QC(WON,SID,PRO,EB,TID,TR)
%UNTITLED Summary of this function goes here
%   WON-Work Order Number; SID-Sample ID; PRO - Process; EB-Employee Badge
%   #; TID-Test Name ID
%   TR- Test Result
%   PF- Pass or Fail

Table=cell(4,6);
for i=1:4
    Table(i, 1)={char(i+64)};
    Table(i, 3)={'Ext'};
end

% Test ID: A, Thickness, ON Extrusion

Table(1, 2)={'Thickness'};
Table(1, 4)={[3.4, 3.4, 4.4, 5.4]};
Table(1, 5)={[3.6, 3.6, 4.6, 5.6]};

% Test ID: B, Dimension, ON Extrusion

Table(2, 2)={'Dimension'};
Table(2, 4)={[119, 89]};
Table(2, 5)={[121, 91]};

% Test ID: C, Peeling, ON Extrusion

Table(3, 2)={'Peeling'};
Table(3, 4)={[50]};
Table(3, 5)={[100000]};

% Test ID: D, Profile Strength, ON Extrusion

Table(4, 2)={'Profile Strength'};
Table(4, 4)={[450]};
Table(4, 5)={[100000]};

TIDD=double(TID)-63;
SIDD=num2str(SID);
SIDDD=str2num(SIDD(3));

LoResult=Table{TIDD, 4}(SIDDD);
HiResult=Table{TIDD, 5}(SIDDD);

if TR < HiResult && TR>LoResult
    P=1;
elseif TR == HiResult
    P=1;
elseif TR == LoResult
    P=1;
else
    P=0;
end
    
[num, txt, ~] = xlsread('QCDB.xlsx');


% Find the Roll for Process
PRROLL=txt(:,1);
t0=strfind(PRROLL,PRO);

[r0,~]=size(t0);
rol=[];
for u=1:r0
    if isempty(t0(u,1))
        rol=[rol, u];
    end
end

%Find the Test (Put text ID (ABCD) on ASTM Column)inside the Process Roll
TIDCOL=txt(:,3);
tididx=[];
for o=rol(1):rol(end)
    tiddd=strfind(TIDCOL(:,o),TID);
    if isempty(tiddd)
        tididx=[tididx, o];
    end
end
%tididx is the roll number for the test to input

% Find the column for the Work Order
WOLINE=txt(1,:);
t1=strfind(WOLINE, WON);

[~, c0]=size(t1);
col=[];
for i=1:c0
    if isempty(t1(1, i))
        col=[col, i];
    end
end

[~, NWO]=size(col);

%Going into each same WO col to see if the test is empty
for j=col(1):col(end)
    %Check if the space is available in sample ID
    if isempty(num(tididx, j))
        %Input Data in this line
        %Input Sample ID
        num(tididx, j)=SID;
        %Input Sample Time
        txt(tididx, j+5)=datestr(now,'mm/dd/yy HH:MM:SS');
        %Input Name
        num(tididx, j+6)=EB;
        %Input Result
        num(tididx, j+4)=TR;
    else
        newcol=j+6;
        while ~isempty(num(tididx,newcol))
            %New Col Found
            newcol=j+6;
            
        end
        %Copy the WO to new Col
        txt(:, newcol:newcol+6)=txt(:, col(1):col(1)+6);
        num(:, newcol:newcol+6)=num(:, col(1):col(1)+6);
        
        
        %Input Data in this line
        %Input Sample ID
        num(tididx, j)=SID;
        %Input Sample Time
        txt(tididx, j+5)=datestr(now,'mm/dd/yy HH:MM:SS');
        %Input Name
        num(tididx, j+6)=EB;
        %Input Result
        num(tididx, j+4)=TR;
    end
end

    


PF = P;
Action = P;
end