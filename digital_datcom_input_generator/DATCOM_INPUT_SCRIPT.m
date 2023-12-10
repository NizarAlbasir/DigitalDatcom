function [] = DATCOM_INPUT_SCRIPT()
[~,~,row]=xlsread('inputExcel.xlsx');

dim=row{1,2};

NMACH=row{2,2};
NALPHA=row{6,2};
NALT=row{4,2};
WT=row{8,2};     LOOP=row{9,2};
SREF=row{10,2};  CBARR=row{11,2};
BLREF=row{12,2};

XCG=row{13,2};   ZCG=row{14,2};
XW=row{15,2};    ZW=row{16,2};
ALIW=row{17,2};  XH=row{18,2};
ZH=row{19,2};    ALIH=row{20,2};
XV=row{21,2};    ZV=row{22,2};
if(row{23,2}==1)
    VERTUP="TRUE";
else
    VERTUP="FALSE";
end


NX=row{24,2};
WDN=row{31,2};
CHRDTP=row{33,2};          SSPN=row{34,2};
SSPNE=row{35,2};           CHRDR=row{36,2};
SAVSI=row{37,2};           CHSTAT=row{38,2};
TYPE=row{40,2};            TWISTA=row{39,2};
DHDADI=row{34,5};


if(row{41,2}==1)
    AIETLP=row{42,2};      NENGSP=row{43,2};
    PHALOC=row{44,2};      PHVLOC=row{45,2};
    PRPRAD=row{46,2}; 
    NOPBPE=row{47,2};
    if(row{48,2}==1)
        CROT="TRUE";
    else
        CROT="FALSE";
    end
end


if(row{49,2}==1)
   AIETLJ=row{50,2};        NENGSJ=row{51,2};
   JEVLOC=row{53,2};        JEALOC=row{52,2};
   JELLOC=row{54,2};        JERAD=row{55,2};
end

HDN=row{56,2};
HCHRDTP=row{58,2};          HSSPN=row{59,2};
HSSPNE=row{60,2};           HCHRDR=row{61,2};
HSAVSI=row{62,2};           HCHSTAT=row{63,2};
HTWISTA=row{64,2};          HTYPE=row{65,2};
HDHDADI=row{59,5};

VDN=row{66,2};
VCHRDTP=row{68,2};          VSSPN=row{69,2};
VSSPNE=row{70,2};
VCHRDR=row{71,2};           VSAVSI=row{72,2};
VCHSTAT=row{73,2};          VTYPE=row{74,2};

FTYPE=row{75,2};
NDELTA=row{76,2};
CHRDFI=row{78,2};           CHRDFO=row{79,2};
SPANFI=row{80,2};           SPANFO=row{81,2};
NTYPE=row{82,2};

STYPE=row{83,2};
ANDELTA=row{84,2};
ACHRDFI=row{86,2};          ACHRDFO=row{87,2};
ASPANFI=row{88,2};          ASPANFO=row{89,2};


file=fopen('input_data.INP','w');
fprintf(file,'DIM %c\n $FLTCON NMACH=%.3f,\n MACH(1)=',dim,NMACH);


for i=2:NMACH+1
    if(mod(i,4)==0)
    fprintf(file,'%.3f,\n ',row{3,i});
    else 
        fprintf(file,'%.3f,',row{3,i});
    end
end

fprintf(file,'\n NALPHA=%.3f,\n ALSCHD(1)=',NALPHA);

for i=2:NALPHA+1
    if(mod(i,4)==0)
    fprintf(file,'%.3f,\n ',row{5,i});
    else 
        fprintf(file,'%.3f,',row{5,i});
    end
end

fprintf(file,'\n NALT=%.3f,\n ALT(1)=',NALT);

for i=2:NALT+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{7,i});
    else 
        fprintf(file,'%.3f,',row{7,i});
    end
end

fprintf(file,'\n WT=%.3f, LOOP=%.1f$\n $OPTINS SREF=%.3f, CBARR=%.3f, BLREF=%.3f$\n ',WT,LOOP,SREF,CBARR,BLREF);
fprintf(file,'$SYNTHS XCG=%.3f, ZCG=%.3f,XW=%.3f, ZW=%.3f,\n ALIW=%.3f, XH=%.3f, ZH=%.3f,\n ALIH=%.3f, XV=%.3f, ZV=%.3f, VERTUP=.%s.$\n $BODY NX=%.1f,\n X(1)=',XCG,ZCG,XW,ZW,ALIW,XH,ZH,ALIH,XV,ZV,VERTUP,NX);

for i=2:NX+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{25,i});
    else 
        fprintf(file,'%.3f,',row{25,i});
    end
end

fprintf(file,'\n R(1)=');

for i=2:NX+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{26,i});
    else 
        fprintf(file,'%.3f,',row{26,i});
    end
end

fprintf(file,'\n ZU(1)=');

for i=2:NX+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{27,i});
    else 
        fprintf(file,'%.3f,',row{27,i});
    end
end

fprintf(file,'\n ZL(1)=');

for i=2:NX+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{28,i});
    else 
        fprintf(file,'%.3f,',row{28,i});
    end
end

fprintf(file,'\n P(1)=');

for i=2:NX+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{29,i});
    else 
        fprintf(file,'%.3f,',row{29,i});
    end
end

fprintf(file,'\n S(1)=');

for i=2:NX+1
    if(mod(i,3)==0 && i==NX+1)
    fprintf(file,'%.3f$',row{30,i});
       else if(mod(i,3)==0 && i~=NX+1)
     fprintf(file,'%.3f,\n ',row{30,i});
       else if(mod(i,3)~=0 && i==NX+1)
        fprintf(file,'%.3f$',row{30,i});
       else
           fprintf(file,'%.3f,',row{30,i});
       end
       end
    end
end

fprintf(file,'\nNACA-W-%d-',WDN);
for i=2:WDN+1
   fprintf(file,'%d',row{32,i});
end

fprintf(file,'\n $WGPLNF CHRDTP=%.3f, SSPN=%.3f, SSPNE=%.3f,\n CHRDR=%.3f, SAVSI=%.3f, CHSTAT=%.2f, TWISTA=%.3f, DHDADI=%.2f, TYPE=%.1f$',CHRDTP,SSPN,SSPNE,CHRDR,SAVSI,CHSTAT,TWISTA,DHDADI,TYPE);

if(row{41,2}==1)
    fprintf(file,'\n $PROPWR AIETLP=%.3f, NENGSP=%.3f, PHALOC=%.3f, PHVLOC=%.3f, PRPRAD=%.3f,',AIETLP,NENGSP,PHALOC,PHVLOC,PRPRAD);
    fprintf(file,'\n NOPBPE=%.3f, CROT=.%s.$\n',NOPBPE,CROT);
end
if(row{49,2}==1)
    fprintf(file,'\n $JETPWR AIETLJ=%.3f, NENGSJ=%.3f, JEVLOC=%.3f, JEALOC=%.3f, JELLOC=%.3f,',AIETLJ,NENGSJ,JEVLOC,JEALOC,JELLOC);
    fprintf(file,'\n JERAD=%.3f$\n',JERAD);
end

fprintf(file,'NACA-H-%d-',HDN);

for i=2:HDN+1
   fprintf(file,'%d',row{57,i});
end

fprintf(file,'\n $HTPLNF CHRDTP=%.3f, SSPN=%.3f, SSPNE=%.3f, TWISTA=%.3f,\n',HCHRDTP,HSSPN,HSSPNE,HTWISTA);
fprintf(file,' CHRDR=%.3f, SAVSI=%.3f, CHSTAT=%.2f, DHDADI=%.2f, TYPE=%.1f$\n',HCHRDR,HSAVSI,HCHSTAT,HDHDADI,HTYPE);
fprintf(file,'NACA-V-%d-',VDN);

for i=2:VDN+1
   fprintf(file,'%d',row{67,i});
end

fprintf(file,'\n $VTPLNF CHRDTP=%.3f, SSPN=%.3f, SSPNE=%.3f,\n',VCHRDTP,VSSPN,VSSPNE);
fprintf(file,' CHRDR=%.3f, SAVSI=%.3f, CHSTAT=%.2f, TYPE=%.1f$\n',VCHRDR,VSAVSI,VCHSTAT,VTYPE);
fprintf(file,'DAMP\nPART\nPLOT\nSAVE\nCASEID WITH ELEVETORS\n');
fprintf(file,' $SYMFLP FTYPE=%.1f, NDELTA=%.1f,\n DELTA(1)=',FTYPE,NDELTA);

for i=2:NDELTA+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{77,i});
    else 
        fprintf(file,'%.3f,',row{77,i});
    end
end

fprintf(file,'\n CHRDFI=%.3f, CHRDFO=%.3f, SPANFI=%.3f, SPANFO=%.3f, NTYPE=%.1f$\nSAVE\nNEXT CASE\n',CHRDFI,CHRDFO,SPANFI,SPANFO,NTYPE);
fprintf(file,' $ASYFLP STYPE=%.1f, NDELTA=%.1f,\n DELTAL(1)=',STYPE,ANDELTA);

for i=2:ANDELTA+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',row{85,i});
    else 
        fprintf(file,'%.3f,',row{85,i});
    end
end

fprintf(file,'\n DELTAR(1)=');


for i=2:ANDELTA+1
    if(mod(i,3)==0)
    fprintf(file,'%.3f,\n ',-row{85,i});
    else 
        fprintf(file,'%.3f,',-row{85,i});
    end
end

fprintf(file,'\n CHRDFI=%.3f, CHRDFO=%.3f, SPANFI=%.3f, SPANFO=%.3f$\nSAVE',ACHRDFI,ACHRDFO,ASPANFI,ASPANFO);

end