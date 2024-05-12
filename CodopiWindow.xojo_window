#tag DesktopWindow
Begin DesktopWindow CodopiWindow
   Backdrop        =   0
   BackgroundColor =   &cFFFFFF
   Composite       =   False
   DefaultLocation =   2
   FullScreen      =   False
   HasBackgroundColor=   False
   HasCloseButton  =   True
   HasFullScreenButton=   False
   HasMaximizeButton=   True
   HasMinimizeButton=   True
   Height          =   400
   ImplicitInstance=   True
   MacProcID       =   0
   MaximumHeight   =   32000
   MaximumWidth    =   32000
   MenuBar         =   256493567
   MenuBarVisible  =   False
   MinimumHeight   =   64
   MinimumWidth    =   64
   Resizeable      =   True
   Title           =   "Untitled"
   Type            =   0
   Visible         =   True
   Width           =   600
End
#tag EndDesktopWindow

#tag WindowCode
	#tag Method, Flags = &h0
		Sub Codopi()
		  Var div, CoLa,test,testmelt,testman,colastart,colaend,colastep,Ctest,ratiofail,MaxPol,chn(15) As double
		  Var dum,dum3,lbyy,domz,domslope,domcept,lig(40),Tpis,Dpis,Psave,SourcePick,Failstring as string
		  Var slope(30),sloperr(30),cept(30),cepterr(30),co(10,16),source(16),sourceRead(16),SourceErr(16),wt(16),wto(16),wtosum As Double
		  var Z,ZlowE,ZhighE,xone,xtwo,ZLow,ZHigh,Zdel,tau,olnum,wtfact,wtsum,zlow1,zlow2,zlow3 as Double
		  Var i,j,k,iCodopi,nlimits,totalcount,DfitPass,TrialIndex,nsstart,olkill,DpxPpx,DoverP,kcode,c1fail,c2fail,c3fail As Integer
		  Var iAl,iCa,iol,iopx,iDeltaD,iOlnum,itcount,firstregressfail,nsSo,nsIo,mmTest(4) As Integer
		  var failmode,StatigPick,dTrial,SourceEl(50) as string
		  Var coord(30),doord(30),minDerror As Integer
		  var OlOpx,DolTol,Dolmax as double
		  Var cpxStart,origcpxstart,cpxEnd,cpxStep,gtStart,origgtstart,gtEnd,gtStep,olStart,olEnd,olStep,ol,ia,ja as double
		  Var uYbLu,uDyEr,uErYb,lYbLu,lDyEr,lErYb  As Double
		  Var deltaD,deltaDmin,CmaxMelt,CoLaError As Double
		  Var iCmaxMelt,killsource,Killup,Killdown,meltkill,gkill,nPass,nRocks,imaxmelt,jmax,jmin as integer 
		  Var iDy,iEr,iYb,iLu,iLa,inSer,inIer,Laposition as integer
		  var clip as clipboard
		  Var sdlg as SaveFileDialog
		  Var Fin as TextInputStream
		  Var frf as TextInputStream
		  Var file,controlfile as FolderItem
		  Var svar,sd(15),sdsum,F(15),Base(15),BaseCh(15),cBase(15),sum,dsum,mean,max,min as double
		  var TestREE,testPC,melts,SFR,SFMM,SCS,MMall,noDP as string
		  var CDPInput,CDPInput1,CDPInput2,CDPInput3,HREEratios as string
		  Var dumrow(30), delim,prelimfails,tname,TPlusName,WtFunction as string
		  Var tdump As FolderItem
		  Var tplusdump As FolderItem
		  var Calculateslope,Cefail,noDPswitch as Boolean
		  Var cdpTest,SaveFailedRatios,SaveFailedMM as Boolean
		  
		  'miscellaneous start settings
		  firstline=0 'allows header to be printed
		  'remtest=false ' turns off NB, the comment above diagram, regr is the variable where NB is stored
		  
		  CDPTail=""
		  FixedName=""
		  prelimfails="CoLa's with failed trials: "
		  WtFunction=""
		  
		  chn(0)=0.237 'La
		  chn(1)= 0.613  'Ce
		  chn(2)=0.0928  'Pr
		  chn(3)=0.457  'Nd
		  chn(4)=0.148  'Sm
		  chn(5)=0.0563  'Eu
		  chn(6)=0.199  'Gd
		  chn(7)=0.0361  'Tb
		  chn(8)=0.246  'Dy
		  chn(9)=0.0546  'Ho
		  chn(10)=0.16  'Er
		  chn(11)=0.0247  'Tm
		  chn(12)=0.161  'Yb
		  chn(13)=0.0246 'Lu
		  
		  nsstart=ns
		  
		  'open CoDoPi controls
		  dom=defltpth + "REEINV"+slash+"CoDoPiControls.txt"
		  controlfile=new FolderItem(dom, FolderItem.PathModes.Shell)
		  frf=TextInputStream.open(controlfile)
		  for i=0 to 27
		    dumrow(i)=frf.readline 'header (0) will be ignored
		  next
		  frf.close
		  
		  'files to read
		  PCFileName = dumrow(1).nthField(chr(9),1)
		  STATIG = dumrow(2).nthField(chr(9),1)
		  SourceName = dumrow(3).nthField(chr(9),1)
		  SAdjustName = dumrow(4).nthField(chr(9),1)
		  HREELimitchoice = dumrow(5).nthField(chr(9),1)
		  
		  'more parameters (6 to 19 in dumrow)
		  for i=1 to 14
		    inputstr(i)=dumrow(i+5).nthField(chr(9),1)
		  next
		  
		  'Boolean parameters and cola traverse range
		  for i=21 to 27' 20-26 in dumrow
		    inputstr(i)=dumrow(i-1).nthField(chr(9),1)
		  next
		  
		  CDPchoices.showmodal 'Fixed vs Traverse etc
		  
		  'clean out any PC info
		  if nmn>0 then
		    for i=1 to nmn 'clear out any existing PC info
		      mn(i)=""
		      if nme>0 then
		        for k=1 to nme
		          lpcf(k) =""
		          ptcf(k, i )=0
		        next
		      end if
		    next
		  end if
		  
		  ' File 1 readPCfile  'load the Partition Coeff data
		  if pcfile<>Nil then
		    'proceed
		  else
		    messagebox "failed to find PC file"
		    exit sub
		  end if
		  
		  dom=pcfile.name
		  PCFileName=dom
		  i =dom.indexOf(".")
		  if i >-1 then
		    dom=dom.left(i)
		  end if
		  CDPName=dom
		  
		  frf=TextInputStream.open(pcfile)
		  
		  dum=frf.readline
		  delim=chr(9)
		  k=dum.CountFields(delim)
		  For i = 2 To k
		    mn(i - 1) = dum.nthField(delim,i)
		  Next
		  
		  nmn = k - 1
		  k = 0
		  while not frf.EndOfFile
		    dum=frf.readline
		    If dum.Length = 0 Then
		      MessageBox "blank line found in file"
		    Else
		      k=k+1
		      lpcf(k) = dum.nthField(delim,1).uppercase 
		      lpcf(k)=labfix(lpcf(k))
		      For i = 2 To nmn + 1
		        ptcf(k, i - 1) = CSingle(dum.nthField(delim,i))
		      Next
		    End If
		  Wend
		  nme = k
		  frf.Close
		  
		  'File 2 Read the Process ID regression data (StatIG)
		  if statigfile<>Nil then
		    'proceed
		  else
		    MessageBox "failed to find Process ID regression data"
		    exit sub
		  end if
		  
		  StatigPick=StatigFile.name
		  Statig=StatigFile.name
		  
		  Fin=TextInputStream.open(StatigFile)
		  dum3=fin.readline
		  domslope=""
		  domz=""
		  domcept=""
		  i = 1
		  
		  While Not Fin.EndOfFile
		    dum3=fin.readline
		    if dum3.left(2)="La" then ' if La present skip and start with Ce
		      dum3=fin.readline
		    end if
		    lbyy=dum3.nthField(chr(9),2)
		    k = lbyy.IndexOf("/")
		    lig(i) = lbyy.Middle(k+1, 2)
		    slope(i)=CSingle(dum3.nthField(chr(9),3))
		    sloperr(i)=CSingle(dum3.nthField(chr(9),4))
		    cept(i)=CSingle(dum3.nthField(chr(9),5))
		    cepterr(i)=CSingle(dum3.nthField(chr(9),6))
		    dum3 = ""
		    
		    If abs(cept(i)) < .25 Then  'find problems -intercepts
		      domcept=domcept+" "+lig(i)
		    end if
		    if cept(i)<=0  then  'zero intercepts
		      domz = domz+" "+lig(i)
		    End If
		    If slope(i) < 0 Then  'find problems -slopes
		      domslope = domslope+" "+lig(i)
		    End If
		    i = i + 1
		  Wend
		  
		  Fin.close
		  iCodopi = i - 1 '# of elements in IG file
		  
		  'for elements with low intercepts, set wt to zero
		  wtosum=0
		  for i=1 to icodopi
		    if abs(cept(i)) < .25 then
		      wto(i)=0
		    else
		      wto(i)=1
		      wtosum=wtosum+1
		    end if
		  next
		  
		  If domz <> "" Then 'Note elements with intercepts < zero
		    MessageBox "Caution "+domz + " has/have intercepts < 0 "
		  End If
		  domz=""
		  If domcept<> "" Then 'Note elements with near zero intercepts
		    MessageBox "Caution "+domcept+ " has/have intercepts near zero and are given zero wt. "+crlf+"These elements will be ignored in the HREE ratio tests"
		  End If
		  domcept=""
		  If domslope <> "" Then 'Note elements with - slopes
		    MessageBox "Caution "+domslope + " have slopes<0 "
		  End If
		  domslope=""
		  
		  For i = 1 To iCodopi 'Ce to last REE from .ig.txt file. Identify positions of HREE used in ratio tests
		    testREE=lig(i)
		    If testREE.compare("Dy")=0 Then
		      iDy=i
		    end if
		    If testREE.compare("Er")=0  Then
		      iEr=i
		    end if
		    If testREE.compare("Yb")=0 Then
		      iYb=i
		    end if
		    If testREE.compare("Lu")=0  Then
		      iLu=i
		    end if
		  next
		  
		  'File 3 Read a source file
		  if SourceFile<>nil then
		    'proceed
		  else
		    MessageBox "failed to find Source file"
		    exit sub
		  end if
		  
		  Sourcename=sourcefile.name
		  frf=TextInputStream.open(sourcefile)
		  dumrow(0)=frf.readline 'header
		  dumrow(1)=frf.readline 'REE values
		  frf.close
		  
		  for i= 0 to 13' this is very restrictive everything has to be in the correct order 14 REE, no PM
		    SourceEl(i)=dumrow(0).nthField(delim,i)
		    If SourceEl(i).compare("La")=0 or SourceEl(i).compare("LA")=0  then
		      LaPosition=i
		      'sgbox str(LaPosition)+"LA position"
		    end if
		  next
		  
		  delim=chr(9) 'a tab
		  SourcePick=Sourcefile.name
		  
		  for i= 0 to 13' this is restrictive everything has to be in the correct order, 14 REE, no PM
		    SourceRead(i)=cdbl(dumrow(1).nthField(delim,LaPosition+2*(i)))
		    SourceErr(i)=cdbl(dumrow(1).nthField(delim,LaPosition+1+2*i)) 'Not currently used in weighting. Source error is largely from Int error, which is used
		  next
		  
		  'File 4  'load preliminary Slope Adjustment estimates
		  dom = defltpth + "REEINV"+slash+SAdjustName
		  file=new FolderItem(dom, FolderItem.PathModes.Shell)
		  if file <>nil then
		    frf=TextInputStream.open(file)
		    if frf<>nil then
		      dumrow(0)=frf.readline 'ignore header
		      k = 1
		      while not frf.EndOfFile
		        dumrow(k)=frf.readline
		        k=k+1
		      wend
		      frf.Close
		      
		      For i = 1 To k-2' for each REE in an IG file, Ce to Lu
		        SAdjust(i)= CSingle(dumrow(i).nthField(delim,2)) 'adjustment to slope
		      next
		      intAdj=CSingle(dumrow(k-1).nthField(delim,2)) 'adjustment to intercept
		      
		    end if
		  end if
		  
		  'File 5 HREE ratio limits for Do profiles
		  'm stands for lower limit
		  'u stands for upper limit
		  dom=defltpth + "REEINV"+slash+"HREE-Limits.txt"
		  file=New FolderItem(dom, FolderItem.PathModes.Shell)
		  if file <>nil then
		    delim=chr(9) 'a tab
		    i=0
		    frf=TextInputStream.open(file)
		    While Not frf.EndOfFile
		      dumrow(i)=frf.readline
		      i=i+1  'Header (0) is ignored
		    wend
		    frf.close
		    nlimits=i-1
		  end if
		  'sgbox hreelimitchoice
		  for k=1 to nlimits
		    dums(k)=dumrow(k).nthField(delim,1)
		    if dums(k).compare(hreelimitchoice) =0 then
		      i=k
		    end if
		  next
		  
		  hreelimitchoice=dumrow(i).nthField(delim,1) 
		  lYbLu= CSingle(dumrow(i).nthField(delim,2))
		  uYbLu= CSingle(dumrow(i).nthField(delim,3))
		  lDyEr= CSingle(dumrow(i).nthField(delim,4))
		  uDyEr=CSingle(dumrow(i).nthField(delim,5))
		  lErYb= CSingle(dumrow(i).nthField(delim,6))
		  uErYb= CSingle(dumrow(i).nthField(delim,7))
		  
		  
		  'Aliasing: match elements from PI regressions - lig() (statigfile) to elements in the partition coeff. file - lpcf()
		  'coord(i) is offset by 1 because there is no La in the lig() data
		  
		  'dom="coord(i)"+crlf
		  For i = 1 To iCodopi 'Ce to Lu, last REE from .ig.txt file
		    testREE=lig(i)
		    For j = 1 To nme  ' matches i = icodopi elements in file.ig and j = nme elements in PC file
		      testPC=lpcf(j)
		      If testREE.compare(testPC)=0 Then
		        coord(i) = j  ' coord(igfile column)= column in PC file
		        'dom=dom+ str(i)+"  "+str(j)+ " " + str(coord(i)) + " " + lig(i) + " " + lpcf(j)+crlf
		      End If
		    Next
		  Next
		  'sgbox dom
		  
		  'dom="Doord(i)"+crlf
		  For i = 1 To iCodopi   ' matches i = icodopi elements in file.ig and j = last elements in main data matrix
		    testREE=lig(i)
		    For j = 1 To last 'elements in main data matrix
		      testPC=l(j)
		      If testREE.compare(testPC)=0 Then
		        doord(i) = j  '    ' doord(igfile column)= column in data file
		        'dom=dom+ str(i) + " " + lig(i) + " " + str(j) + " " + l(j)+" " +str(doord(i))+crlf
		      End If
		    Next
		    'sgbox l(doord(i))+"  elems "+str(doord(i))
		  Next
		  
		  iLa=0
		  
		  For i = 1 To last 'find La in main matrix
		    If l(i).compare("La")=0  Then
		      iLa=i
		    end if
		  next
		  
		  j=0
		  For i = ila To ila+28 'find error in main matrix
		    If l(i).indexof("error")>-1  Then
		      j=j+1
		    end if
		  next
		  
		  If j>7 then
		    'skip
		  else
		    dom="Here are the REE data columns" +crlf
		    For i = ila To ila+14
		      dom=dom+l(i)+crlf
		    next
		    dom=dom+crlf
		    dom=dom+"If error columns, such as errorLa, are absent, then error cannot be propagated."+crlf
		    dom=dom+"To easily add empty error columns to your REE data, do the following."+crlf
		    dom =dom+"Read into Igpet:  Master_Corrected_REEs+ERROR.txt, from the IgpetDocs/REEINV folder."+crlf
		    dom=dom+ "Use Igpet's ADD( Merge) function (File menu) to add your data file."+crlf 
		    dom=dom+ "This adds empty error columns to allow propagation"+crlf
		    dom=dom+ "Plot the result in a REE diagram, then click the SaveSelect Button"+crlf
		    dom=dom+ "Edit the name and use this file for the inverse."
		    messagebox dom
		  end if
		  
		  lig(0)="La"
		  
		  'column names for Fixed calculations, the u(row,column) matrix
		  l(last+0) = "CoLa"
		  l(last+1) = "P"+mn(1) 'pGt
		  l(last+2) = "P"+mn(2)
		  l(last+3) = "P"+mn(3)
		  l(last+4) = "P"+mn(4)
		  l(last+5) = "P"+mn(3)+"+"+mn(4)
		  l(last+6) = "Pol#"
		  l(last+7) = "D"+mn(1)
		  l(last+8) = "D"+mn(2)
		  l(last+9) = "D"+mn(3)
		  l(last+10) = "D"+mn(4)
		  l(last+11) = "D"+mn(3)+"+"+mn(4)
		  l(last+12) = "Dol#"
		  iOlnum=last+12
		  l(last+13) = "MaxMelt"
		  imaxmelt=last+13
		  l(last+14) =  "D[fit]"
		  ideltaD=last+14
		  l(last+15) = "Pol-Limit"
		  l(last+16) = "%F min"
		  l(last+17) = "%F max"
		  l(last+18) = "D[o] Dy[N]/Er[N]"
		  l(last+19) = "D[o] Er[N]/Yb[N]"
		  l(last+20) = "D[o] Yb[N]/Lu[N]"
		  l(last+21) = "TrialIndex"
		  '22new columns
		  
		  'SETUP Choices for CoLa and Meltmode minerals
		  dums(0) = "Setup choices"
		  dums(1) = "CoLa" 'for fixed file
		  
		  dums(2) = "Start of Pgt loop (0)"
		  dums(3) = "End of Pgt  loop (15)"
		  dums(4) = "Step of Pgt  loop (1)"
		  
		  dums(5) = "Start of Pcpx loop"
		  dums(6) = "End of Pcpx loop"
		  dums(7) = "Step of Pcpx loop"
		  
		  dums(8) = "Start of Pol# loop (0)"
		  dums(9) = "End of Pol# loop (100)"
		  dums(10)= "Step of Pol# loop (5)"
		  
		  dums(11) = "Dol# tolerance"
		  dums(12) = "Dol# high"
		  dums(13) = "Max OL in melt"
		  dums(14) = "CoLa error in %"
		  inputstr(15)=mstr(lYbLu)
		  dums(15) = "Lowest Yb/Lu"
		  inputstr(16)=mstr(uYbLu)
		  dums(16) = "Highest Yb/Lu"
		  inputstr(17)=mstr(lDyEr)
		  dums(17) = "Lowest Dy/Er"
		  inputstr(18)=mstr(uDyEr)
		  dums(18) = "Highest Dy/Er"
		  inputstr(19)=mstr(lErYb)
		  dums(19) = "Lowest Er/Yb"
		  inputstr(20)=mstr(uErYb)
		  dums(20) = "Highest Er/Yb"
		  dums(21) = "Save failed mantle modes (Fixed)"  '21 to 24 are true/false questions, boolean
		  dums(22) = "Save failed melt modes (Fixed)" 
		  dums(23) = "Save ion rad vs f(i) (Fixed)"
		  dums(24)="Use DCe>PCe test"
		  dums(25) = "CoLa start"
		  dums(26) = "CoLa stop"
		  dums(27) = "CoLa step"
		  if imode=2 or imode=3 then
		    nbuts = 27
		  else
		    nbuts = 24
		  end if
		  Inputfrm.Showmodal
		  If dom="QUIT" Then
		    exit sub
		  end if
		  
		  CoLa=CSingle(inputstr(1))
		  gtStart= CSingle(inputstr(2))
		  origgtstart=gtstart ' fossil
		  gtEnd = CSingle(inputstr(3))
		  gtStep = CSingle(inputstr(4))
		  cpxStart= CSingle(inputstr(5))
		  origcpxstart=cpxstart
		  cpxEnd =CSingle(inputstr(6))
		  cpxStep = CSingle(inputstr(7))
		  olStart= CSingle(inputstr(8))
		  olEnd =CSingle(inputstr(9))
		  olStep = CSingle(inputstr(10))
		  DolTol=CSingle(inputstr(11))
		  Dolmax=CSingle(inputstr(12))
		  MaxPol=CSingle(inputstr(13))
		  CoLaError=CSingle(inputstr(14))
		  lYbLu= CSingle(inputstr(15))
		  uYbLu= CSingle(inputstr(16))
		  lDyEr= CSingle(inputstr(17))
		  uDyEr= CSingle(inputstr(18))
		  lErYb= CSingle(inputstr(19))
		  uErYb= CSingle(inputstr(20))
		  HREEratios="lYbLu="+inputstr(15)+" uYbLu="+inputstr(16)+crlf+"'"
		  HREEratios=HREEratios+"lDyEr="+ inputstr(17)+" uDyEr="+ inputstr(18)+crlf+"'"
		  HREEratios=HREEratios+"lErYb="+ inputstr(19)+" uErYb="+inputstr(20)
		  
		  SFMM=inputstr(21)'saveFailedMantleModes"   'inputstr(21)
		  SaveFailedMM=Boolean.FromString(SFMM)
		  SFR=inputstr(22)
		  SaveFailedRatios=Boolean.FromString(SFR)
		  SCS=inputstr(23)
		  Calculateslope=Boolean.FromString(SCS)
		  noDP=inputstr(24)
		  noDPswitch=Boolean.FromString(noDP)
		  
		  if imode=2 or imode=3 then
		    colastart=CSingle(inputstr(25))
		    colaend=CSingle(inputstr(26))
		    colastep=CSingle(inputstr(27))
		  end if
		  
		  CDPInput1="pGt~ "+inputstr(2)+" "+inputstr(3)+" "+inputstr(4)+" pCpx~ "+inputstr(5)+" "+inputstr(6)+" "+inputstr(7)+" pOl~ "+inputstr(8)+" "+inputstr(9)+" "+inputstr(10)
		  CDPInput2=" Dol~ "+inputstr(11)+" "+inputstr(12)
		  CDPInput3=" CoLa~ "+inputstr(25)+" "+inputstr(26)+" "+inputstr(27)
		  CDPinput=CDPInput1+CDPInput2+CDPInput3
		  
		  for i=1 to nmn
		    If mn(i).uppercase.indexOf("GT")>-1 then  ' find garnet from PC file 
		      iAl=i
		    end if
		    If mn(i).uppercase.indexOf("CPX")>-1 then 'find cpx  from PC file 
		      iCa=i
		    end if
		    if mn(i).uppercase.indexOf("OL")>-1 then 'fiind ol
		      iol=i
		    end if
		    if mn(i).uppercase.indexOf("OPX")>-1 then   'find opx
		      iopx=i
		    end if
		  next
		  
		  for i=1 to ns 'load some fakes into matrix for samples so Subselect won't throw them out
		    u(i,last+18)=".75"  ' fake  Dy/Er for Do  for all the melts and parents etc.
		    u(i,last+19)=".80" ' fake  Er/Yb for Do  for all the melts and parents etc.
		    u(i,last+20)=".95" '  fake  Yb/Lu for Do  for all the melts and parents etc.
		    u(i,last+1)="1"  ' fake  pGt  for all the melts and parents etc.
		    u(i,last+6)="1"  ' fake  dGt 
		  next
		  
		  if imode=2 or imode=3 then '2 means traverse mode
		    tname="traverse~"+cdpinput+"~"+StatigPick
		  end if
		  
		  if imode=2 or imode=3 then 
		    tdump= SpecialFolder.Desktop.Child(tname)
		    
		    sdlg=new saveFileDialog
		    sdlg.InitialFolder =SpecialFolder.Desktop
		    sdlg.SuggestedFileName =Tdump.name
		    sdlg.promptText="Select/edit a name for Traverse output"
		    sdlg.filter="text/plain"
		    tdump=sdlg.showmodal()
		    
		    if tdump<>nil then
		      'proceed
		    else
		      exit sub
		    end if
		    
		    if imode=3 then
		      TPlusName="P-"+Tdump.name
		      tplusdump= SpecialFolder.Desktop.Child(TPlusName)
		      
		      sdlg=new saveFileDialog
		      sdlg.InitialFolder =SpecialFolder.Desktop
		      sdlg.SuggestedFileName =TPlusName
		      sdlg.promptText="Select/edit a name for Traverse-Plus output"
		      sdlg.filter="text/plain"
		      Tplusdump=sdlg.showmodal()
		      
		      if tplusdump<>nil then
		        'proceed
		      else
		        exit sub
		      end if
		    end if
		    
		    FixedName=""
		  else
		    'Fixed CoLa setting
		    colastart=cola 'to fake a loop when doing just one CoLa
		    colaend =cola+0.1
		    colastep=1
		    FixedName="Fixed at~"+str(cola)+" ~"+cdpinput+"~"+StatigPick
		  end if
		  
		  'ReWrite CodopiControls here"
		  dom=defltpth + "REEINV"+slash+"CoDoPiControls.txt"
		  file=New FolderItem(dom, FolderItem.PathModes.Shell)
		  Var frout as TextOutputStream
		  
		  if file <>nil then
		    delim=chr(9) 'a tab
		    i=0
		    frout=TextOutputStream.Create(file)
		    frout.writeline("Value"+chr(9)+"Description")
		    
		    frout.writeline PCFileName
		    frout.writeline statig
		    frout.writeline SourceName
		    frout.writeline SAdjustName+chr(9)+"Initial Slope-Int adjustments"
		    frout.writeline hreelimitchoice+chr(9)+"HREE ratio limits"
		    
		    frout.writeline str(CoLa)+chr(9)+"CoLa, chondritic units"
		    frout.writeline str(gtstart)+chr(9)+"gtstart"
		    frout.writeline str(gtend)+chr(9)+"gtend"
		    frout.writeline str(gtStep)+chr(9)+"gtstep"
		    frout.writeline str(cpxstart)+chr(9)+"cpxstart"
		    frout.writeline str(cpxend)+chr(9)+"cpxend"
		    frout.writeline str(cpxStep)+chr(9)+"cpxstep"
		    frout.writeline str(olstart)+chr(9)+"ol#start"
		    frout.writeline str(olend)+chr(9)+"ol#end"
		    frout.writeline str(olStep)+chr(9)+"ol#step"
		    frout.writeline str(DolTol)+chr(9)+"Dol# tolerance"
		    frout.writeline str(Dolmax)+chr(9)+"Dol# high"
		    frout.writeline str(MaxPol)+chr(9)+"Max %olivine in melt mode"
		    frout.writeline str(CoLaError)+chr(9)+"CoLa error in %" 
		    frout.writeline SaveFailedMM.toString+chr(9)+"Save failed mantle modes (Fixed option)"
		    frout.writeline SaveFailedRatios.toString+chr(9)+"Save failed melt modes (Fixed option)"
		    frout.writeline Calculateslope.tostring+chr(9)+"Save ion.rad(i) vs f(i) (Fixed option)"
		    frout.writeline noDPswitch.tostring+chr(9)+"Use DCe>PCe test"
		    'frout.writeline str(wtfact) +chr(9)+"Wt factor, 0 set att wts to 1"
		    frout.writeline str(colastart)+chr(9)+"colastart"
		    frout.writeline str(colaend)+chr(9)+"colaend"
		    frout.writeline str(colastep)+chr(9)+"colastep"
		    frout.close
		  end if
		  
		  'sgbox "set up the Error Propagation here. THIS IS A FOSSIL except for the message"
		  div=1000
		  k=0
		  sdsum=0
		  for i=1 to icodopi
		    svar=pow(sloperr(i)/slope(i),2)+pow(cepterr(i)/cept(i),2) '(slope int adjustments ignored, they should cancel out) This is %std err
		    sd(i)=pow(svar,0.5)  ' this Standard Dev. is for error propagation
		    if sd(i)<div then
		      div=sd(i) 'find the minimum
		    end if
		    if sd(i)<.00001 then
		      k=k+1
		    end if
		    sdsum=sdsum+sd(i)
		  next
		  
		  
		  if k>5 then 
		    MessageBox mstr(k)+" REE with extremly low SDs, probably a synthetic model with zero error."+crlf+" Weights set to 1.0 "
		    wtfact=0
		  end if
		  wtfact=0' kills weighting function
		  
		  dom="Weighting Function for synthetic models from IG- data"+crlf
		  
		  for i=1 to icodopi
		    wt(i)=wto(i)*(1-wtfact*(sd(i)*13/sdsum)) 'wtfact=0 above so this is a fossil
		    dom=dom+lig(i)+"   "+format(sd(i)*13/sdsum, "-#.##")'+format(wt(i), "-#.###")+" "
		    'f wt(i)<.5 then 
		    'dom=dom+" Low wt!"
		    'j=j+1
		    'end if
		    dom=dom+crlf
		  next
		  'If wtfact>0 then
		  'msgbox dom
		  dom=""
		  'end if
		  
		  dums(0)="Eliminate elements by setting weight to 0"
		  for i=1 to icodopi
		    dums(i) = lig(i)
		    if wt(i)<.5 then 
		      dums(i)=dums(i)+" Low Intercept"
		    end if
		    inputstr(i)=format(wt(i), "-#.###")
		  next
		  
		  nbuts =icodopi
		  Inputfrm.Showmodal
		  
		  dom="Final Weighting Function"+crlf'"  WTfactor=" +str(wtfact)+"  "+crlf
		  
		  for i=1 to icodopi
		    dums(i) = lig(i)
		    wt(i)=csingle(inputstr(i))
		    dom=dom+"'"+lig(i)+"   "+format(wt(i), "-#.###")+crlf
		  next
		  WtFunction=dom
		  
		  'Make rough calculations of melt proportions for each sample to pick maximum and minimum samples
		  'Melt% depend on the CoLa value selected at the top of the Setup choices
		  'Tm, Yb, Lu typically have large errors and are ignored in picking max and min. La is also not part of this.
		  
		  max=0 'These are important
		  min=1  
		  jmax=0
		  jmin=0
		  
		  if calculateslope=true then ' melt% pattern slopes (Fixed mode) should be horizontalish if batch melting but some are not suggesting outliers
		    'Column labels for the melt proportions for each sample if true 
		    l(last) = "Intercept"
		    l(last+1) = "Slope"
		    l(last+2)="mean melt proportion"
		    l(last+3)="CoLa"
		  end if
		  
		  
		  'Locate jmax and jmin, positions of max and min samples
		  dom=""
		  For k= 1 to nfound
		    j = vn(k)
		    sum=0
		    wtsum=0
		    for i=1 to icodopi-3 'Ce to Er. (La=0. Tm, Yb, Lu typically have large errors)
		      div=csingle(u(j,  doord(i)))/chn(i)
		      
		      if div.Equals(0, 1) or cept(i).Equals(0, 1) then 'this test allows blank lines to pass through
		        'skip
		      else
		        f(i)=(CoLa/cept(i))*((1/div)-SAdjust(i)*slope(i)) 'verified 2-24-21
		        u(ns+j, doord(i))=str(f(i))
		        if f(i)<0 or f(i).Equals(0, 1) then
		        else
		          sum=sum+wto(i)*f(i)
		          wtsum=wtsum+wto(i)
		        end if
		        
		        If calculateslope=true then ' melt% pattern slopes
		          ap(i, 1) = REErad(i)
		          ap(i, 2) = f(i)
		        end if
		      end if
		    next 'i REE
		    
		    mean=sum/wtsum
		    u(ns+j, 2)=str(mean)' put ave in La position
		    if mean>max then
		      max=mean
		      jmax=j
		    end if
		    
		    if mean<min then
		      min=mean
		      jmin=j
		    end if
		    
		    s(ns +j) = "melt estimate for " +s(j)
		    u(ns+j, isam)=s(ns+j)
		    ki(ns+j)=13
		    g(ns+j)=13
		    u(ns+j, ik)="13" 'blue hexagon
		    symscale(ns+j)=1
		    vn(nfound+j) = ns + j -nfound
		    
		    if calculateslope=true then ' melt% pattern slopes
		      polynom(10, 99) 'looks at first 10 REE, 99 means first order poly.
		      u(ns+j, last)=format(polyslope, "-#.#####")
		      u(ns+j, last+1)=format(polyint, "-#.#####")
		      u(ns+j, last+2)=format(mean, "-#.#####")
		      u(ns+j, last+3)=format(cola, "-##.##")
		    end if
		    
		  next 'k samples
		  
		  
		  if imode=1 then 'Fixed mode only
		    melts="Min-Max Melts% are "+format(100*min, "##.##")+"    "+format(100*max, "##.##") 
		    messagebox "maximum melt= "+s(jmax)+"    ~"+format(100*max,"-##.# ")+crlf+"minimum melt= "+s(jmin)+"   ~ " +format(100*min,"-##.#")+crlf+"Click OK and then wait"
		  end if
		  
		  For i = nfound+1 To ns 
		    vn(i) = ns + i 
		  Next
		  
		  nfound=nfound+ns' This is for fixed mode. Line 804 resets ns to nsstart for traverse mode
		  ns=nfound
		  
		  If calculateslope=true then'  calculate linear fit to ionic radius vs melt% and then exit Codopi
		    App.MouseCursor = System.Cursors.StandardPointer
		    last=last+4
		    dom= "Melt proportion data and slopes of melt% vs ionic radius have been calculated and added to data in memory"
		    dom=dom+"Reread your data file to proceed"+crlf
		    dom=dom+crlf+"Exiting inverse routine"
		    messagebox dom
		    exit sub
		  end if
		  
		  App.MouseCursor = System.Cursors.Wait
		  
		  'START CoLa loop
		  
		  ratioFail=Ctest
		  colaend =colaend + 0.05*colastep ' to include the last one
		  
		  'increase the ends slightly so the last one gets included
		  gtend=gtend+.05*gtstep
		  cpxend=cpxend+.05*cpxstep
		  olend=olend+.05*olstep
		  
		  
		  For Ctest=colastart to colaend step colastep
		    TrialIndex=0
		    cola=ctest
		    
		    minDerror=0 'an index
		    'S fake placeholders
		    'I
		    
		    'SOURCE as a function of CoLa
		    for j=0 to 13 'sources data from file must emanate from La[N]=1 (0.237ppm). Cola raises or lowers the initial source profile
		      source(j+1)=SourceRead(j)*CoLa ' the +1 puts elements in correct order for the SIadjust subroutine
		    next
		    
		    if imode= 2 or imode=3  then
		      ns=nsstart 'for each step of the Traverse, this sets ns to the orig # of smples in the REE plot
		    end if
		    
		    'Top of data field (jmin) based on an average wt% based on Ce-Er
		    'For the initial step, SAdjust(i) is only approximate and the results for the first step may be erratic. So put in a couple of lower steps that can be ignored.
		    'SAdjust(i) is recalculated multiple times so its value becomes reliable in just an iteration or two
		    'The Top of the sample field using sample(jmin)
		    
		    sum=0
		    k=0
		    for i=1 to icodopi-3 'Ce to Er (La=0, Tm, Yb and Lu typically have large errors)
		      div=csingle(u(jmin,  doord(i)))/chn(i)
		      
		      if div.Equals(0, 1) or cept(i).Equals(0, 1) then
		        'skip
		      else
		        f(i)=(CoLa/cept(i))*((1/div)-SAdjust(i)*slope(i)) 'verified 2-24-21
		        if f(i).Equals(0, 1) or f(i)<0 then
		          'skip
		        else
		          sum=sum+wto(i)*f(i)
		          k=k+wto(i)
		        end if
		        
		      end if
		    next 'i  REE
		    min=sum/k
		    
		    for i=1 to icodopi
		      baseCh(i)=CoLa/(min*IntAdj*cept(i)+SAdjust(i)*slope(i)*CoLa)
		      base(i)=baseCh(i)*chn(i) 'ppm REE
		      u(ns+1,doord(i))=str(base(i)) 'actually Top
		    next 'i  REE
		    
		    
		    'Fake melt modes in F top row
		    u(ns + 1, last+1) =mstr(5) ' melt modes (P)
		    u(ns + 1, last+2) =mstr(20)
		    u(ns + 1, last+3) =mstr(40)
		    u(ns + 1, last+4) =mstr(35)
		    u(ns + 1, last+5) =mstr(75)
		    u(ns + 1, last+6) =mstr(35/75)
		    
		    nfound=nfound+1
		    ns=ns+1
		    s(ns) = "Top"
		    u(ns, isam)="Top"
		    ki(ns)=14
		    g(ns)=14
		    u(ns, ik)="14"  'filled blue hex
		    symscale(ns)=1
		    
		    'Base of data field (jmax), based on an average F% based on Ce-Er
		    'For the initial step, SAdjust(i) is only approximate and the results for the first step may be erratic. So put in a couple of lower steps that can be ignored.
		    'SAdjust(i) is recalculated multiple times so its value becomes reliable in just an iteration or two
		    'Base uses sample(jmax)
		    
		    'The base is done after Top because the vaules in base(i) are needed later for the Source test.
		    sum=0
		    k=0
		    for i=1 to icodopi 'Ce to Lu. (La=0)
		      div=csingle(u(jmax,  doord(i)))/chn(i)
		      if div.Equals(0, 1) or cept(i).Equals(0, 1) then
		        'skip
		      else
		        f(i)=(CoLa/cept(i))*((1/div)-SAdjust(i)*slope(i)) 'verified 2-24-21
		        if f(i).Equals(0, 1) or f(i)<0 then
		          'skip
		        else
		          sum=sum+wto(i)*f(i)
		          k=k+wto(i)
		        end if
		      end if
		    next 'i  REE
		    max=sum/k
		    'sgbox str(Cola)+"  in cola loop  "+ str(max)
		    
		    for i=1 to icodopi' using all the REE but La here
		      baseCh(i)=CoLa/(max*IntAdj*cept(i)+SAdjust(i)*slope(i)*CoLa)
		      base(i)=baseCh(i)*chn(i)
		      u(ns+1,doord(i))=str(base(i))
		    next 'i  REE
		    
		    u(ns + 1, last+22)=format(100*max,"-00.##")
		    'Fake melt modes in F base row
		    u(ns + 1, last+1) =mstr(5) ' melt modes (P)
		    u(ns + 1, last+2) =mstr(20)
		    u(ns + 1, last+3) =mstr(40)
		    u(ns + 1, last+4) =mstr(35)
		    u(ns + 1, last+5) =mstr(75)
		    u(ns + 1, last+6) =mstr(35/75)
		    
		    nfound=nfound+1
		    ns=ns+1
		    s(ns) = "Base"
		    u(ns, isam)="Base"
		    ki(ns)=14
		    g(ns)=14
		    u(ns, ik)="14"
		    symscale(ns)=1
		    
		    'save slope and intercept vectors from Statig
		    for i=1 to icodopi' using all the REE but La here
		      u(ns + 1,doord(i))=str(slope(i))
		      u(ns + 1,doord(i)+1)=str(sloperr(i))
		    next 'i  REE
		    
		    nfound=nfound+1
		    ns=ns+1
		    s(ns) = "S[o] vector"
		    u(ns, isam)=s(ns)
		    ki(ns)=20
		    g(ns)=20
		    u(ns, ik)="20" ' symbol key
		    symscale(ns)=1
		    nsSo=ns
		    
		    for i=1 to icodopi' using all the REE but La here
		      u(ns + 1,doord(i))=str(cept(i))
		      u(ns + 1,doord(i)+1)=str(cepterr(i))
		    next 'i  REE
		    nfound=nfound+1
		    ns=ns+1
		    nsIo=ns
		    s(ns) = "I[o] vector"
		    u(ns, isam)=s(ns)
		    ki(ns)=20
		    g(ns)=20
		    u(ns, ik)="20"
		    symscale(ns)=1
		    
		    nrocks=ns    'Important: # of samples plus Top+Base+S(i)+I(i)
		    
		    'Prep for Main Melt Mode Loop
		    deltaDmin=10^6
		    dTrial=""
		    
		    totalcount=0
		    killsource=0
		    killup=0
		    killdown=0
		    olkill=0
		    gkill=0
		    DoverP=0
		    DpxPpx=0
		    FirstregressFail= 0
		    
		    'test for dimension warning
		    test=(1+(gtEnd-gtStart)/gtStep)*(1+(cpxEnd-cpxStart)/cpxStep)*(1+(olEnd-olStart)/olStep)
		    If test >1000+irw and ctest.equals(colastart,1) then
		      Messagebox "Your grid has "+str(test)+" gridpoints." +crlf+"If the program crashes, make some steps larger "+crlf+"to lower the number."
		    end if
		    
		    for i=1 to 4
		      mmtest(i)=0
		    next
		    
		    c1fail=0
		    c2fail=0
		    c3fail=0
		    
		    'MELT MODE LOOPS start here
		    
		    For ia = gtStart To gtEnd step gtStep  'gt
		      meltmode(1) = ia  'gt
		      For ja=cpxStart to cpxEnd step cpxStep 
		        meltmode(2)=ja ' cpx
		        OlOpx=100-meltmode(1)-meltmode(2) 'Ol+Opx by difference
		        
		        if OlOpx<0 then
		          OlOpx=0
		        end if
		        
		        for ol=olStart to olEnd step olstep
		          meltmode(3)=OlOpx*ol/100                   'ol
		          meltmode(4)=100-ia-ja-meltmode(3)   'opx
		          
		          totalcount=totalcount+1
		          
		          tpis ="P:"+mstr(meltmode(1))+"-"+mstr(meltmode(2))+"-"+mstr(meltmode(3))+"-"+mstr(meltmode(4))+"-Pol#="+mstr(meltmode(3)/(meltmode(3)+meltmode(4)))'tpis is a label
		          Psave=mstr(meltmode(1))+chr(9)+mstr(meltmode(2))+chr(9)+mstr(meltmode(3))+chr(9)+mstr(meltmode(4)) +chr(9) 'this is data
		          Psave=Psave+mstr(meltmode(4)+meltmode(3)) +chr(9)+mstr(100*meltmode(3)/(meltmode(3)+meltmode(4)))
		          'dom="Pi calcs"+crlf
		          For j = 1 To icodopi ' just get P's can't use icodopi because it lacks La
		            i=coord(j) 'CORRECT
		            pm(i) = 0
		            pm(i) = pm(i) + ptcf(i, iAl) * meltmode(1)/ 100 
		            pm(i) = pm(i) + ptcf(i, iCa) * meltmode(2)/ 100 
		            pm(i) = pm(i) + ptcf(i, iOl) * meltmode(3)/ 100 
		            pm(i) = pm(i) + ptcf(i, iOpx)* meltmode(4)/ 100 
		          Next
		          
		          cdpTest=True
		          
		          For i = 1 To iCodopi ' for each REE in IG file, Ce to Lu
		            co(1, i) = pm(coord(i))    'Pi's
		            
		            if cept(i).Equals(0,1) then
		              cept(i)=.00001
		            end if
		            co(2, i) = (1 - pm(coord(i))) / (IntAdj*cept(i))  'Co's   Co
		            div=pow(cepterr(i)/cept(i),2)+pow(CoLaError/100,2)
		            div=pow(div,0.5)  ' this is % standard error. is for error propagation
		            If div.Equals(0,1) then
		              div = 0.00001
		            End If
		            co(3, i)=div*co(2, i)' Co error is 1 std error
		            
		            ' adjusted Do
		            co(5, i) =SAdjust(i)*slope(i) * co(2, i)' Do's adjusted for La=0 but without chondrite norm factor or CoLa
		            div=pow(sloperr(i)/slope(i),2)+pow(cepterr(i)/cept(i),2)+pow(CoLaError/100,2) '(slope and int adjustments ignored, they cancel out) This is %variance
		            div=pow(div,0.5)  ' this is % standard error. is for error propagation
		            co(7, i)=div*co(5, i)' error 1 std error
		          Next  'i
		          
		          if meltmode(3)>MaxPol then
		            olkill=olkill+1
		            failmode="MAXOL" ' >Max ol in melt"
		            cdptest=false
		          else
		            
		            ' the HREE ratio tests and source test may be premature because we have not decided on the adjustment to Slope
		            if uDyEr<co(5, iDy) / co(5, iEr) and wto(8).Equals(1, 1) and wto(10).Equals(1, 1) then
		              cdptest=false
		              failmode="Up"
		              Killup=killup+1
		            elseif uErYb<co(5, iEr) / co(5, iYb) and wto(10).Equals(1, 1) and wto(12).Equals(1, 1) then
		              cdptest=false
		              failmode="Up"
		              Killup=killup+1
		            elseif uYbLu<co(5, iYb)/ co(5, iLu) and wto(12).Equals(1, 1)and wto(13).Equals(1, 1) then
		              cdptest=false
		              failmode="Up"
		              Killup=killup+1
		            elseif lDyEr>co(5, iDy) / co(5, iEr) and wto(8).Equals(1, 1) and wto(10).Equals(1, 1) then
		              cdptest=false
		              failmode="Down"
		              killdown=killdown+1
		            elseif  lErYb>co(5, iEr) / co(5, iYb) and wto(10).Equals(1, 1) and wto(12).Equals(1, 1) then
		              cdptest=false
		              failmode="Down"
		              killdown=killdown+1
		            elseif  lYbLu>co(5, iYb)/ co(5, iLu)and wto(12).Equals(1, 1) and wto(13).Equals(1, 1) then
		              cdptest=false
		              failmode="Down"
		              killdown=killdown+1
		              
		              'Source test (Dy,Er,Yb,Lu) assumes that the errors from using preliminary SAdjust values to determine f(i) and Base(i) are trivial 
		            elseif CoLa*co(2, iDy)>base(iDy)/chn(iDy) and wto(8).Equals(1, 1) then
		              cdptest=false
		              killsource=killsource+1
		              failmode="Source"
		            elseif CoLa*co(2,iEr)>base(iEr)/chn(iEr) and wto(10).Equals(1, 1) then
		              cdptest=false
		              killsource=killsource+1
		              failmode="Source"
		            elseif CoLa* co(2, iYb)>base(iYb)/chn(iYb) and wto(12).Equals(1, 1) then
		              cdptest=false
		              killsource=killsource+1
		              failmode="Source"
		            elseif CoLa*co(2, iLu) >base(iLu)/chn(iLu) and wto(13).Equals(1, 1) then
		              cdptest=false
		              killsource=killsource+1
		              failmode="Source"
		            end if
		          end if
		          
		          if cdptest=false then 'the trial failed 
		            
		            if ratiofail.Equals(ctest, 1)then 'shows all the trials with some failure
		              'skip
		            else
		              prelimfails=prelimfails+str(ctest)+" "
		              ratiofail=ctest
		            end if
		            
		            if SaveFailedRatios=True then
		              
		              select case failmode
		              case "UP"
		                kcode=27
		              case "DOWN"
		                kcode=28
		              case "SOURCE"
		                kcode=9
		              case "MAXOL"
		                kcode=26
		              end select
		              
		              trialindex=trialindex+1
		              s(ns+1) = "Pi " + tpis+"    "+str(trialindex)
		              u(ns+1,isam)=s(ns + 1)
		              ki(ns+1) = kcode
		              u(ns+1,ik)=str(kcode)
		              g(ns+1) = kcode
		              u(ns + 1, last+21) =str(TrialIndex)
		              
		              u(ns + 1, last) =mstr(CoLa) 
		              u(ns + 1, last+1) =mstr(meltmode(1)) 
		              u(ns + 1, last+2) =mstr(meltmode(2))
		              u(ns + 1, last+3) =mstr(meltmode(3))
		              u(ns + 1, last+4) =mstr(meltmode(4))
		              u(ns + 1, last+5) =mstr(meltmode(3)+meltmode(4))
		              u(ns + 1, last+6) =mstr(100*meltmode(3)/(meltmode(3)+meltmode(4)))
		              
		              For i= 1 To  icodopi 'Do - error not determined
		                div=co(5, i) * chn(i)  'chn(i)  normalization converts from normalized values to measured values. 
		                div=div*CoLa            ' here is where the initial value of La is used (CoLa)
		                u(ns + 1, doord(i)) = format(div,"-##.#######")
		                div=co(7, i) * chn(i)
		                div=div*CoLa   'error
		                u(ns + 1, doord(i)+1) = format(div,"-##.#######") 'adds error
		              next
		              u(ns + 1, ila)="0" 'no La  May3 2024
		              
		              symscale(ns+1)=1
		              vn(ns+1)=ns+1
		              ns=ns+1
		              nfound=nfound+1
		            end if
		            
		          else
		            'Trials that pass
		            TrialIndex=TrialIndex+1
		            
		            'i add melt modes to Pi(1), Co(2), Do(3), Dc(4), Sc(5), Ic(6) last two are dubious
		            for i=1 to 4 '4 not 6
		              u(ns + i, last) =mstr(CoLa) 
		              u(ns + i, last+1) =mstr(meltmode(1)) 
		              u(ns + i, last+2) =mstr(meltmode(2))
		              u(ns + i, last+3) =mstr(meltmode(3))
		              u(ns + i, last+4) =mstr(meltmode(4))
		              u(ns + i, last+5) =mstr(meltmode(3)+meltmode(4))
		              u(ns + i, last+6) =mstr(100*meltmode(3)/(meltmode(3)+meltmode(4)))
		            next
		            
		            'Begin golden section to find olnum
		            
		            ZLowE=0
		            zlow1=0
		            zlow2=0
		            zlow3=0
		            
		            tau= (pow(5,.5)-1)/2 'could push this constant further up out of loops
		            ZLow=0
		            ZHigh=Dolmax '100
		            
		            xone=Zlow+(1-tau)*(Zhigh-Zlow)' 0.382
		            xtwo=Zlow+tau*(Zhigh-Zlow)'        0.618
		            
		            Zdel=Zhigh-Zlow
		            itcount=0
		            
		            While Zdel>DolTol 
		              
		              Z=xone
		              for j=1 to icodopi
		                i=coord(j)
		                ac(j, 1) = wt(j)*ptcf(i,1)
		                ac(j, 2) = wt(j)*ptcf(i,2)
		                ac(j, 3) = wt(j)*(Z*ptcf(i,3)+ ptcf(i,4)*(100-Z))/100
		                ac(j, 4) = wt(j)*CoLa* SAdjust(j)*slope(j) * (1 - pm(i)) / (IntAdj*cept(j))    'CoLa*Do
		              next
		              ac(icodopi+1,1)=1 'gt place  sum mantle mineral proportions to 1
		              ac(icodopi+1,2)=1 'cpx place
		              ac(icodopi+1,3)=1 'ol+opx place
		              ac(icodopi+1,4)=1 'this is needed, it is c4 in these eqns:  c1*GT(i)+c2*CPX(i)+c3*OLOPX(i)=c4
		              
		              calcs(icodopi+1,3)' Z=xone  in wend loop
		              
		              mantmode(1)=100*cf(1) ' mantle modes (D) 
		              mantmode(2)=100*cf(2)
		              mantmode(3)=100*cf(3)*Z/100
		              mantmode(4)=100*cf(3)*(100-Z)/100
		              
		              For j= 1 To  icodopi '
		                i=coord(j)
		                Dm(i) = 0 'sum set to zero
		                Dm(i) = Dm(i) + ptcf(i, iAl) * cf(1)
		                Dm(i) = Dm(i) + ptcf(i, iCa) * cf(2)'TempCPX
		                Dm(i) = Dm(i) + ptcf(i, iOl) * cf(3)*Z/100'TempOL
		                Dm(i) = Dm(i) + ptcf(i, iOpx)* cf(3)*(100-Z)/100'TempOPX
		                'dom=dom+lpcf(i)+" "+ str(olnum)+"   "+str(i)+chr(9)+str(ptcf(i,ial))+"  "+str(Dm(i))+crlf
		              Next
		              
		              ZlowE=0
		              For j=1 to icodopi
		                i=coord(j)
		                deltaD=wto(j)*(Dm(i)-wt(j)*CoLa* SAdjust(j)*slope(j) * (1 - pm(i)) / (IntAdj*cept(j)))
		                ZlowE=ZlowE+deltaD*deltaD' squared delta
		              next
		              
		              Z=xtwo
		              
		              for j=1 to icodopi
		                i=coord(j)
		                ac(j, 1) = wt(j)*ptcf(i,1)
		                ac(j, 2) = wt(j)*ptcf(i,2)
		                ac(j, 3) = wt(j)*(Z*ptcf(i,3)+ ptcf(i,4)*(100-Z))/100
		                ac(j, 4)=wt(j)*CoLa* SAdjust(j)*slope(j) * (1 - pm(i)) / (IntAdj*cept(j))    'CoLa*Do
		              next
		              
		              ac(icodopi+1,1)=1 'gt place  sum mantle mineral proportions to 1
		              ac(icodopi+1,2)=1 'cpx place
		              ac(icodopi+1,3)=1 'ol+opx place
		              ac(icodopi+1,4)=1 'this is needed, it is c4 in these eqns:  c1*GT(i)+c2*CPX(i)+c3*OLOPX(i)=c4
		              
		              calcs(icodopi+1,3)' Z=xtwo  in wend loop
		              
		              mantmode(1)=100*cf(1) ' mantle modes (D) 
		              mantmode(2)=100*cf(2)
		              mantmode(3)=100*cf(3)*Z/100
		              mantmode(4)=100*cf(3)*(100-Z)/100
		              
		              For j= 1 To  icodopi '
		                i=coord(j)
		                Dm(i) = 0 'sum set to zero
		                Dm(i) = Dm(i) + ptcf(i, iAl) * cf(1)
		                Dm(i) = Dm(i) + ptcf(i, iCa) * cf(2)'TempCPX
		                Dm(i) = Dm(i) + ptcf(i, iOl) * cf(3)*Z/100'TempOL
		                Dm(i) = Dm(i) + ptcf(i, iOpx)* cf(3)*(100-Z)/100'TempOPX
		                'dom=dom+lpcf(i)+" "+ str(olnum)+"   "+str(i)+chr(9)+str(ptcf(i,ial))+"  "+str(Dm(i))+crlf
		              Next
		              ZhighE=0
		              For j=1 to icodopi
		                i=coord(j)
		                deltaD=wto(j)*Dm(i)-wt(j)*CoLa* SAdjust(j)*slope(j) * (1 - pm(i)) / (IntAdj*cept(j))  
		                ZhighE=ZhighE+deltaD*deltaD' squared delta
		              next
		              
		              If ZLowE<ZhighE then
		                ZHigh=xtwo
		                'Tdom=Tdom+"Flush high"
		              else
		                ZLow=xone
		                'Tdom=Tdom+"Flush low"
		              end if
		              
		              itcount=itcount+1
		              Zdel=abs(xtwo-xone)' ok this is the item to test to pass wend
		              
		              xone=Zlow+(1-tau)*(Zhigh-Zlow) 'Golden section reset
		              xtwo=Zlow+tau*(Zhigh-Zlow)
		              
		              olnum=(xone+xtwo)/2
		              
		              if itcount>10 or Zdel.Equals(0, 1)  then
		                messagebox "Failed to converge to limit"
		                Zdel=0 'exits While - wend
		              end if
		              
		              SIadjust(meltmode(),mantmode(),source())
		              'end if
		            wend
		            
		            'Now use best ol opx to make a model to save
		            'make one last adjustment
		            
		            for j=1 to icodopi
		              i=coord(j)
		              ac(j, 1) = wt(j)*ptcf(i,1)
		              ac(j, 2) = wt(j)*ptcf(i,2)
		              ac(j, 3) = wt(j)*(olnum*ptcf(i,3)+ ptcf(i,4)*(100-olnum))/100
		              ac(j, 4)=wt(j)*CoLa* SAdjust(j)*slope(j) * (1 - pm(i)) / (IntAdj*cept(j))    'CoLa*Do
		            next
		            
		            ac(icodopi+1,1)=1 'gt place in sum mantle proportions to 1
		            ac(icodopi+1,2)=1 'cpx place
		            ac(icodopi+1,3)=1 'ol+opx place
		            ac(icodopi+1,4)=1 'this is needed, it is c4 in these eqns:  c1*GT(i)+c2*CPX(i)+c3*OLOPX(i)=c4
		            calcs(icodopi+1,3) 'olnum decided so a real calc
		            
		            if cf(1)<0 or cf(2)<0 or cf(3)<0 then
		              
		              'golden section failures' Pi values saved here (negative mantle minerals)
		              
		              s(ns+1) = "Pi " + tpis+"  Failstring= "+failstring +"    "+str(trialindex)
		              u(ns+1,isam)=s(ns + 1)
		              u(ns+1,ik)=mstr(17)
		              ki(ns + 1) =17
		              g(ns + 1) =17
		              u(ns + 1, last+21) =str(TrialIndex)
		              if cf(1)<0 then
		                c1fail=c1fail+1
		              else
		                if cf(2)<0 then
		                  c2fail=c2fail+1
		                else
		                  if cf(3)<0 then
		                    c3fail=c3fail+1
		                  end if
		                end if
		              end if
		              
		              if SaveFailedMM=True and imode=1 then
		                
		                'Put actual REE values for golden fails into the data matrix
		                s(ns+2) = "Co-CDP-" + str(CoLa)+" "+tpis+dpis+"    "+str(TrialIndex)
		                u(ns+2,isam)=s(ns + 2) 
		                ki(ns+2) = 21
		                u(ns+2,ik)="21"
		                g(ns +2) =21
		                
		                s(ns+3) = "Do "+ str(CoLa)+" " + tpis +dpis+"    "+str(TrialIndex)
		                u(ns+3,isam)=s(ns + 3)  'for saving
		                ri(ns+3) = TrialIndex 'jcode
		                ki(ns+3) =33              'open green cross, kcode symbol. all remaining trials start as a pass "3"
		                u(ns+3,ik)="33"       'kcode to save
		                g(ns+3) = 33             'code of symbol to be plotted
		                
		                'Put actual REE values for golden fails into the data matrix, 3 lines per trial
		                
		                For i= 1 To  icodopi 'Co
		                  div=co(2, i) * chn(i)  'chn(i) normalization converts from normalized values to measured values. 
		                  div=div*CoLa               ' here is where the initial value of La is used (CoLa)
		                  u(ns + 2, doord(i)) = format(div,"-##.#######")'   Co vector
		                  div=co(3, i) * chn(i)
		                  div=div*CoLa   
		                  u(ns + 2, doord(i)+1) = format(div,"-##.#######")  'adds Co error
		                next
		                symscale(ns+2)=1
		                
		                'Add La to Co profiles
		                u(ns + 2, ila)= mstr(CoLa*chn(0)) 'La is located at 0
		                
		                For i= 1 To  icodopi 'Do
		                  div=co(5, i) * chn(i)  'chn(i)  normalization converts from normalized values to measured values. 
		                  div=div*CoLa               ' here is where the initial value of La is used (CoLa)
		                  u(ns + 3, doord(i)) = format(div,"-##.#######")'    Do vector
		                  div=co(7, i) * chn(i)
		                  div=div*CoLa   'error
		                  u(ns + 3, doord(i)+1) = format(div,"-##.#######") 'adds Do error
		                next
		                symscale(ns+3)=1
		                
		                For i= 1 To  icodopi 'Pi Ce to Lu
		                  div=co(1, i) * chn(i)  'chn(i) normalization converts from normalized values to measured values. 
		                  u(ns + 1, doord(i)) = format(div,"-##.#######")  'Pi
		                  u(ns + 1, doord(i)+1) = "0" 'Pi has no error
		                next
		                symscale(ns+1)=1
		                u(ns + 1, ila)="0" 'no La
		                
		                dom=""
		                
		                vn(nfound+1) = ns + 3 -nfound
		                ns=ns +3
		                nfound = nfound + 3
		              else
		                vn(nfound+1) = ns + 1 -nfound
		                ns=ns +1
		                nfound = nfound + 1
		              end if
		              
		              
		              
		            else
		              
		              mantmode(1)=100*cf(1) ' mantle modes (D) 
		              mantmode(2)=100*cf(2)
		              mantmode(3) =100*cf(3)*olnum/100
		              mantmode(4) =100*cf(3)*(100-olnum)/100
		              SIadjust(meltmode(),mantmode(),source())
		              
		              'for j=2 to icodopi+1
		              'nslope(j-1)=SIReturn(j,1)
		              'ncept(j-1)=SIReturn(j,2)
		              'next
		              
		              for i= 1 to icodopi
		                if cept(i).Equals(0,1) then
		                  cept(i)=.00001
		                end if
		                co(2, i) = (1 - pm(coord(i))) / (IntAdj*cept(i))  'Co's   Co
		                If div.Equals(0,1) then
		                  div = 0.00001
		                End If
		                div=pow(cepterr(i)/cept(i),2)+pow(CoLaError/100,2)
		                div=pow(div,0.5)  ' this is % standard error for error propagation
		                If div.Equals(0,1) then
		                  div = 0.00001
		                End If
		                co(3, i)=div*co(2, i)' error 1 std error
		                
		                co(5, i)=SAdjust(i)*slope(i) * co(2, i)'  Do
		                div=pow(sloperr(i)/slope(i),2)+pow(cepterr(i)/cept(i),2)+pow(CoLaError/100,2) '(slope and int adjustments ignored, they should cancel out) This is %variance
		                div=pow(div,0.5)  ' this is % standard error for error propagation
		                co(7, i)=div*co(5, i)' error 1 std error
		              next
		              'end if
		              
		              s(ns+1) = "Pi " + tpis+"    "+str(TrialIndex)
		              u(ns + 1, last+21) =str(TrialIndex)
		              dpis=" D: "+format(mantmode(1),"-##.##")+"gt " +format(mantmode(2),"-##.##")+"cpx " +format(mantmode(3),"-##.##")+"ol "+format(mantmode(4),"-##.##")+"opx"
		              
		              'Determine max melt
		              if cf(1)<0 or cf(2)<0 or cf(3)<0 then  
		                CmaxMelt=2*max
		              else
		                iCmaxMelt=0
		                CmaxMelt=1
		                For i = 1 To 4
		                  if meltmode(i).Equals(0,1) then
		                    'skips zero values for meltmode
		                  else
		                    test=mantmode(i)/meltmode(i)
		                    
		                    if test<CmaxMelt then
		                      CmaxMelt=test
		                      iCmaxmelt=i
		                    end if
		                  end if
		                next
		                
		                
		              end if
		              
		              'REAL stuff
		              u(ns + 1, last+21) =str(TrialIndex)
		              u(ns + 2, last+21) =str(TrialIndex)
		              u(ns + 3, last+21) =str(TrialIndex)
		              
		              
		              s(ns+1) = "Pi " + tpis+"    "+str(TrialIndex)
		              u(ns+1,isam)=s(ns + 1)
		              ki(ns+1) = 1
		              u(ns+1,ik)="1" 'kcode to save
		              g(ns+1) = 1
		              
		              for i=1 to 4'6
		                u(ns+i,iolnum)=format(olnum,"###.##")
		              next
		              
		              s(ns+2) = "Co-CDP-" + str(CoLa)+" "+tpis+dpis+"    "+str(TrialIndex)
		              u(ns+2,isam)=s(ns + 2) 
		              ki(ns+2) = 2 '  blue circle
		              u(ns+2,ik)="2"  'kcode to save
		              g(ns +2) = 2
		              
		              s(ns+3) = "Do "+ str(CoLa)+" " + tpis +dpis+"    "+str(TrialIndex)
		              u(ns+3,isam)=s(ns + 3)  'for saving
		              ri(ns+3) = TrialIndex 'jcode
		              ki(ns+3) =3              'filled green circle kcode symbol all remaining trials start as a pass 
		              u(ns+3,ik)="3"         'kcode to save
		              g(ns+3) = 3             'code of symbol to be plotted
		              
		              s(ns+4) = "Dc "  + dpis
		              u(ns+4,isam)=s(ns + 4) 
		              ki(ns+4) = 4
		              u(ns+4,ik)="4"
		              g(ns + 4) = 4
		              
		            end if
		            
		            'calculate Dm from mantle mode
		            For j= 1 To  icodopi
		              i=coord(j)
		              Dm(i) = 0 'sum set to zero
		              Dm(i) = Dm(i) + ptcf(i, iAl) * mantmode(1)/100
		              Dm(i) = Dm(i) + ptcf(i, iCa) * mantmode(2)/100
		              Dm(i) = Dm(i) + ptcf(i, iOl) * mantmode(3)/100
		              Dm(i) = Dm(i) + ptcf(i, iOpx)* mantmode(4)/100
		              'dom=dom+format(olnum,"##.#")+"   "+str(i)+"  "+str(ptcf(i,ial))+"  "+str(Dm(i))+"  "+crlf
		            Next
		            
		            'Calculate Base from max
		            For j= 1 To  icodopi
		              i=coord(j)
		              cBase(j) = cola*Co(2,j)/(Dm(i)+max*(1-Pm(i)))
		              baseCh(i)=CoLa/(max*IntAdj*cept(i)+SAdjust(i)*slope(i)*CoLa)
		            Next
		            
		            'Here the actual REE values go into the data matrix, 4 lines per trial
		            
		            For i= 1 To  icodopi 'Co
		              div=co(2, i) * chn(i)  'chn(i) normalization converts from normalized values to measured values. 
		              div=div*CoLa               ' here is where the initial value of La is used (CoLa)
		              u(ns + 2, doord(i)) = format(div,"-##.#######")'   Co vector
		              div=co(3, i) * chn(i)
		              div=div*CoLa   
		              u(ns + 2, doord(i)+1) = format(div,"-##.#######")  'adds Co error
		            next
		            symscale(ns+2)=1
		            
		            'Add La to Co profiles
		            u(ns + 2, ila)= mstr(CoLa*chn(0)) 'La is located at the 0 position
		            
		            For i= 1 To  icodopi 'Do
		              div=co(5, i) * chn(i)  'chn(i)  normalization converts from normalized values to measured values. 
		              div=div*CoLa               ' here is where the initial value of La is used (CoLa)
		              u(ns + 3, doord(i)) = format(div,"-##.#######")'    Do vector
		              div=co(7, i) * chn(i)
		              div=div*CoLa   'error
		              u(ns + 3, doord(i)+1) = format(div,"-##.#######") 'adds Do error
		            next
		            symscale(ns+3)=1
		            
		            For i= 1 To  icodopi 'Pi        Ce to Lu
		              div=co(1, i) * chn(i)  'chn(i) normalization converts from normalized values to measured values. 
		              u(ns + 1, doord(i)) = format(div,"-##.#######")  'Pi
		              u(ns + 1, ila)= "0" 'no La
		            next
		            symscale(ns+1)=1
		            u(ns + 1, ila)="0" 'La is located at the 0 position
		            dom=""
		            
		            For j= 1 To  icodopi 'Dc
		              i=coord(j)
		              u(ns + 4, doord(j))=str(Dm(i)* chn(j) ) 'Dc added to matrix
		            Next
		            symscale(ns+4)=1
		            
		            
		            'now throw out trials with flaws
		            
		            'First test, all Do symbols are 3 to start with
		            if u(ns+3,ik)=mstr(3) then
		              
		              if CmaxMelt<max then         'Max melt fail 12
		                for i=1 to 4
		                  u(ns+i,ik)=mstr(0) 
		                  ki(ns + i) =0
		                  g(ns + i) =0
		                next
		                u(ns+3,ik)=mstr(12) 
		                ki(ns + 3) = 12
		                g(ns + 3) = 12
		                
		                mmtest(iCmaxmelt)=mmtest(iCmaxmelt)+1
		              end if
		            end if
		            
		            'no Dcpx> Pcpx  
		            if u(ns+3,ik)=mstr(3) then
		              'If noDPswitch=true then
		              If Mantmode(2)>meltmode(2) then' Do will be > Pi     2=cpx
		                for i=1 to 4 
		                  u(ns+i,ik)=mstr(0) 
		                  ki(ns + i) =0
		                  g(ns + i) =0
		                next
		                u(ns+3,ik)=mstr(11) 
		                ki(ns + 3) = 11
		                g(ns + 3) = 11
		              end if
		              'end if
		            end if
		            
		            'Test Ce.Do (eqn8)-Ce.Pi-input
		            if u(ns+3,ik)=mstr(3) then ' much better test range
		              If noDPswitch=true then
		                Cefail=false
		                div= u(ns + 3,4).ToDouble'  Do for Ce  3=Do, 4=Ce
		                test=u(ns + 1,4).ToDouble'  Pi for Ce    1=Pi,  4=Ce
		                
		                if div>test then' Ce do>pi
		                  'fail
		                  Cefail=true
		                end if
		                
		                if Cefail=true then         'Cefail 10
		                  for i=1 to 4 
		                    u(ns+i,ik)=mstr(0) 
		                    ki(ns + i) =0
		                    g(ns + i) =0
		                  next
		                  u(ns+3,ik)=mstr(10) 'This overrides previous value of 3
		                  ki(ns + 3) = 10
		                  g(ns + 3) = 10
		                end if
		              end if
		            end if
		            
		            if cf(1)<0 or cf(2)<0 or cf(3)<0 then   'skip fit tests if negative mantle minerals
		              'skip
		            else
		              'calculate the fit tests
		              
		              sum=0 'Dfit
		              'dsum=0 'Abs(Dif)
		              
		              For j=1 to icodopi
		                i=coord(j)
		                deltaD=wto(j)*(Dm(i)-CoLa*co(5,j))
		                'dsum=dsum+Abs(deltaD)/cola
		                sum=sum+deltaD*deltaD/(cola)^2
		                
		              next
		              
		              if u(ns+3,ik)="3" then ' failed trials are not used to revise minimum
		                
		                if sum<deltaDmin  then
		                  deltaDmin=sum
		                  minDerror=ns+3' 1 is P, 2 is Co, 3 is Do, 4 is Dc,      obsolete 5 is Sc, 6 is Ic
		                  dTrial=str(TrialIndex)
		                end if
		              end if
		              
		              'P    Co    err  -err    Do  err  -err  Dm
		              for i=1 to 4 '4 not 6
		                u(ns + i, last+7) =mstr(mantmode(1)) ' mantle modes (D) put into melt mode (P) row (1) for comparisons
		                u(ns + i, last+8) =mstr(mantmode(2))
		                u(ns + i, last+9) =mstr(mantmode(3))
		                u(ns + i, last+10) =mstr(mantmode(4))
		                u(ns + i, last+11) =mstr(mantmode(3)+mantmode(4)) 'ol+opx
		                u(ns + i, last+12) =mstr(100*mantmode(3)/(mantmode(3)+mantmode(4)))
		                u(ns + i, last+13) =format(100*CmaxMelt,"-##.##")
		                u(ns + i, last+14)=format(sum,"-0.00000e")
		                'u(ns + i, last+15)=format(dsum,"-0.00000e")
		                u(ns + i,last+15)=mstr(MaxPol)
		                u(ns + i,last+16)=format(100*min,"-##.##")
		                u(ns + i,last+17)=format(100*max,"-##.##")
		              next
		              
		              for i= 1 to 4 'loads up HREE ratios
		                if co(5, iEr).Equals(0,1) then
		                  u(ns + i, last+18)=""
		                else
		                  u(ns + i, last+18) =mstr(co(5, iDy) / co(5, iEr))
		                end if
		                if co(5, iYb).Equals(0,1) then
		                  u(ns + i, last+19) =""
		                else
		                  u(ns + i, last+19) =mstr(co(5, iEr) / co(5, iYb))
		                end if
		                if co(5, iLu).Equals(0,1) then
		                  u(ns + i, last+20) =""
		                else
		                  u(ns +i, last+20) =mstr(co(5, iYb) / co(5, iLu))
		                end if
		              next
		              
		              
		              if nfound<0 then
		                nfound=0
		              end if
		              For i = nfound + 1 To nfound +4  '6
		                vn(i) = ns + i -nfound
		              Next
		              ns=ns +4  '6
		              nfound = nfound +4  '6
		              
		            end if 'last
		          end if ' fail cdp test
		        next 'ol
		      next 'cpx
		    next 'gt
		    
		    ' U matrix should be complete by this point
		    If imode>1 then
		      nbol(0)=-90 'turns off symbols in Symform (Traverse)
		      nbol(4)=-90
		      nbol(8)=-90
		      nbol(9)=-90
		      nbol(10)=-90
		      nbol(11)=-90
		      nbol(12)=-90
		      nbol(13)=-90
		      nbol(14)=-90
		      nbol(17)=-90
		      nbol(18)=-90
		      nbol(19)=-90
		      nbol(20)=-90
		      nbol(21)=-90
		      nbol(22)=-90
		      nbol(26)=-90
		      nbol(27)=-90
		      nbol(28)=-90
		    else
		      nbol(0)=-90 'turns off symbols in Symform (Fixed)
		      nbol(1)=-90 
		      nbol(2)=-90 
		      nbol(4)=-90
		      nbol(8)=-90
		      nbol(9)=-90
		      
		      nbol(13)=-90
		      nbol(14)=-90
		      
		      nbol(18)=-90
		      nbol(19)=-90
		      nbol(20)=-90
		      
		      nbol(22)=-90
		      nbol(26)=-90
		      nbol(27)=-90
		      nbol(28)=-90
		    end if
		    
		    
		    DfitPass=0 'get ready to count errors etc
		    meltKill=0
		    DpxPpx=0
		    DoverP=0
		    
		    For i = nrocks+1  To ns   'Ignore first 4 entries after samples 
		      
		      if u(i,ik)="3"  then
		        DfitPass=DfitPass+1
		      end if
		      
		      if u(i,ik)="10" then
		        DoverP=DoverP+1
		        symscale(i)=1
		      end if
		      
		      if u(i,ik)="11" then ' noDcpx>Pcpx
		        DpxPpx=DpxPpx+1
		        symscale(i)=1
		      end if
		      
		      if u(i,ik)="12" then
		        meltkill=meltkill+1
		        symscale(i)=1
		      end if
		      iCmaxmelt=0
		      
		      if u(i,ik)="17" then 'counting a Pi but it is a failed mantle mode so Pi is all there is.
		        gkill=gkill+1
		        symscale(i)=1
		      end if
		      
		    next
		    
		    nfound=0
		    j=1
		    For i = 1 To ns 'This loop eliminates the symbols for 0, Dm, Sr, Sc, Ir, Ic and error counts
		      if g(i)=0 or g(i)=4 or g(i)=8 or g(i)=18 or g(i)=19 or g(i)=20 or g(i)=10 or g(i)=11 or g(i)=12 or g(i)=13 or g(i)=26 or g(i)=27 or g(i)=28 or g(i)=17  then
		        'skip
		      else
		        vn(j) = i 
		        j=j+1
		        nfound=nfound+1
		      end if
		    Next
		    'sgbox str(ns)+"=ns  "+str(nfound)+ "=nfound "+ str(itc)+"=itc"
		    
		    App.MouseCursor = System.Cursors.StandardPointer
		    
		    if imode=1 then 'fixed Cola
		      
		      dom="'Fixed CoLa = "+mstr(CoLa)+crlf+"'"
		      dom=dom+Wholefilename+crlf+"'"
		      dom=dom+statigpick+crlf+"'"
		      dom=dom+sourcepick+crlf+"'"
		      dom=dom+PCfile.name+" partition coeffcients"+crlf+"'"
		      dom=dom+SAdjustname+"  Initial SI-adjustment"+crlf+"'"
		      dom=dom+hreelimitchoice+"  HREE ratio pick"+crlf+"'"
		      dom=dom+HREEratios+crlf+"'"
		      dom=dom+mstr(MaxPol)+" MaxPol"+crlf+"'"
		      dom=dom+"Switches"+crlf+"'"
		      dom=dom+SFMM+"    Save failed mantle modes (Fixed)" +crlf+"'" 
		      dom=dom+SFR+"    Save failed melt modes (Fixed)" +crlf+"'" 
		      dom=dom+SCS+"    Save ion rad vs f(i) (Fixed)"+crlf+"'" 
		      dom=dom+MMall+"    Add ol,opx to Maxmelt calc."+crlf+"'" '    MMall
		      dom=dom+noDP+"    Use DCe>PCex test"+crlf+"'" 
		      dom=dom+"Search limits"+crlf+"'"
		      dom=dom+CDPinput1+crlf+"'"
		      dom=dom+CDPinput2+crlf+"'"
		      dom=dom+CDPinput3+crlf+"'"
		      dom=dom+melts+crlf+"'"
		      dom=dom+"Trials= " +mstr(totalcount) + crlf +"'"
		      dom=dom+"Too much Ol in meltmode= "+mstr(olkill)+crlf +"'"
		      dom=dom+"HREE ratios < min= "+mstr(killdown)+crlf +"'"
		      dom=dom+"HREE ratios > max= "+mstr(killup)+crlf +"'"
		      dom=dom+"Source too high=  " +mstr(killsource )+crlf+"'"
		      dom=dom+"Remainder = "+mstr(totalcount-killdown-killup-olkill-killsource)+crlf+"'"
		      dom=dom+"GoldenSection failed to find Dol#=  " +mstr(gkill)+crlf+"'"
		      dom=dom+"  (-) garnet count=  " +mstr(c1fail)+crlf+"'"
		      dom=dom+"  (-) cpx count=  " +mstr(c2fail)+crlf+"'"
		      dom=dom+"  (-) ol+opx count=  " +mstr(c3fail)+crlf+"'"
		      dom=dom+"CmaxMelt fails = " + mstr(meltKill) +crlf+"'"
		      dom=dom+"fGt="+mstr(mmtest(1))+" fCpx="+mstr(mmTest(2))+" fOl="+mstr(mmTest(3))+" fOpx="+mstr(mmTest(4))+crlf+"'"
		      dom=dom+"Dcpx>Pcpx  " +mstr(DpxPpx )+crlf+"'"
		      dom=dom+"D-Ce>P-Ce=  " +mstr(DoverP)+crlf+"'"
		      'dom=dom+"No OPX in mantle=  " +mstr(NoOPX )+crlf+"'"
		      dom=dom+"Total Passes = "+mstr(nPass+DfitPass)+crlf+"'"
		      
		      CDPTail="Stop"+crlf+dom+"'"+WtFunction
		      
		      clip=New Clipboard
		      clip.Text=dom
		      clip.Close
		      dom=dom+"The message just displayed is now on the Clipboard."+crlf+crlf
		      messagebox dom
		      
		      dom="Calculations are not automatically saved"+chr(10)+"Use the File/SaveFile option in the menu." +chr(10)+"A name will be suggested."+crlf+crlf
		      dom=dom +crlf + "To run a different model, reread the geochem data and replot the spiderdiagram."+crlf
		      messagebox dom
		    end if
		    
		    if imode=1 then
		      yyb = 0.001
		      yye = 100
		      nspremod=ns   'HELP Nov11 brought this back to make Histograms work
		      zerosin = True
		      'sgbox str(last)
		      last=last+22 ' this is needed for Fixed mode
		      exit sub
		    end if
		    
		    if imode=2 then
		      dom="'Traverse = "+crlf+"'"
		    end if
		    if imode=3 then
		      dom="'TraversePlus = "+crlf+"'"
		    end if
		    
		    if imode=2 or imode=3 then ' Traverse
		      dom=dom+Wholefilename+crlf+"'"
		      dom=dom+statigpick+crlf+"'"
		      dom=dom+sourcepick+crlf+"'"
		      dom=dom+PCfile.name+" partition coeffcients"+crlf+"'"
		      dom=dom+SAdjustname+"  Initial SI-adjustment"+crlf+"'"
		      dom=dom+hreelimitchoice+"  HREE ratio pick"+crlf+"'"
		      dom=dom+HREEratios+crlf+"'"
		      dom=dom+mstr(MaxPol)+" MaxPol"+crlf+"'"
		      dom=dom+"Switches"+crlf+"'"
		      'dom=dom+SFMM+"    Save failed mantle modes (Fixed)" +crlf+"'" 
		      'dom=dom+SFR+"    Save failed melt modes (Fixed)" +crlf+"'" 
		      'dom=dom+SCS+"    Save ion rad vs f(i) (Fixed)"+crlf+"'" 
		      dom=dom+MMall+"    Add ol,opx to Maxmelt calc."+crlf+"'" '    MMall
		      dom=dom+noDP+"    Use DCe>PCe test"+crlf+"'" 
		      dom=dom+"Search limits"+crlf+"'"
		      dom=dom+CDPinput1+crlf+"'"
		      dom=dom+CDPinput2+crlf+"'"
		      dom=dom+CDPinput3+crlf+"'"
		      'dom=dom+"Wt function factor, 0 means equal= "+str(wtfact)+crlf+"'"
		      CDPTail="Stop"+crlf+dom+WtFunction
		      
		      if firstline=0 then
		        dom="sample"+chr(9)+"Kcode"+chr(9)+"Description"+chr(9)+"CoLa"+chr(9)+"Pgt"+chr(9)+"Pcpx"+chr(9)+"Pol"+chr(9)+"Popx"+chr(9)+"Pol+Popx"+chr(9)+"Pol#"+chr(9)
		        dom=dom+"Dgt"+chr(9)+"Dcpx"+chr(9)+"Dol"+chr(9)+"Dopx"+chr(9)+"Dol+Dopx"+chr(9)+"Dol#"+chr(9)+"MeltMax"+chr(9)+"D[fit]"+chr(9)
		        dom=dom+"Pol-Limit"+chr(9)
		        dom=dom+"%F min"+chr(9)+"%F max"+chr(9)+"DyEr"+chr(9)+"ErYb"+chr(9)+"YbLu"+chr(9)
		        dom=dom+"Trials"+chr(9)+"too much Ol in melt"+chr(9)+"HREE-fail-down"+chr(9)+"HREE-fail-up"+chr(9)
		        dom=dom+"Source-fails"+chr(9)+"Remainder"+chr(9)
		        dom=dom+"Dol# fail"+chr(9)+"(-) garnet"+chr(9)+"(-) cpx"+chr(9)+"(-) ol+opx"+chr(9)
		        dom=dom+"FailMaxMelt"+chr(9)+"fGt"+chr(9)+"fCpx"+chr(9)+"fOl"+chr(9)+"fOpx"+chr(9)
		        dom=dom+"Dcpx>Pcpx"+chr(9)+"D-Ce>P-Ce"+chr(9)+"Passes"+crlf 
		        ' adding lines at start of file to describe what the input was. For cases that crash these lines are clues.
		        dom=dom+"'Traverse  "+crlf+"'"
		        dom=dom+Wholefilename+crlf+"'"
		        dom=dom+statigpick+crlf+"'"
		        dom=dom+sourcepick+crlf+"'"
		        dom=dom+cdpinput+crlf
		        if imode=2 then
		          firstline=firstline+1
		        end if
		      else
		        dom=""
		      end if
		      
		      'sgbox str(minderror)
		      i=minDerror'     Put data for minimum Dfit into traverse
		      dom=dom+mstr(CoLa)+"  "+dTrial+chr(9)+u(i,ik)+chr(9)+CDPinput+chr(9)+mstr(CoLa)+chr(9)
		      
		      for j= last+1 to last+20  
		        dom=dom+u(i,j)+chr(9)
		      next
		      dom=dom+mstr(totalcount)+chr(9)+mstr(olkill)+chr(9)+mstr(killdown)+chr(9)+mstr(killup)+chr(9)
		      dom=dom+mstr(killsource)+chr(9)+mstr(totalcount-killdown-killup-olkill-killsource-firstregressfail)+chr(9)
		      dom=dom+mstr(gkill)+chr(9)+mstr(c1fail)+chr(9)+mstr(c2fail)+chr(9)+mstr(c3fail)+chr(9)
		      dom=dom+mstr(meltKill)+chr(9)+mstr(mmtest(1))+chr(9)+mstr(mmTest(2))+chr(9)+mstr(mmTest(3))+chr(9)+mstr(mmTest(4))+chr(9)
		      dom=dom+mstr(DpxPpx)+chr(9)+mstr(DoverP)+chr(9)+mstr(DfitPass)
		      
		      Var output As TextOutputStream
		      Try
		        output = TextOutputStream.Open(tdump)
		        output.WriteLine dom
		        output.Close
		      Catch e As IOException
		        messagebox "Unable to append to file."
		      End Try
		      
		      if imode=3 then 'create/add to a file with all (4) profiles for the selected best fit for each CoLa value
		        Var Poutput As TextOutputStream
		        Try
		          Poutput = TextOutputStream.Open(tplusdump)
		          if firstline=0 then
		            dom=l(0)+chr(9)+"Kcode"
		            For j=ila to last+17
		              dom=dom+chr(9)+l(j)
		            next 
		            Poutput.Writeline dom
		            
		            dom=u(nsSo,0)+ chr(9)+"20"
		            For j=ila to last+17
		              dom=dom+chr(9)+u(nsSo,j)
		            next 
		            Poutput.Writeline dom
		            
		            dom=u(nsIo,0)+ chr(9)+"20"
		            For j=ila to last+17
		              dom=dom+chr(9)+u(nsIo,j)
		            next 
		            Poutput.Writeline dom
		            
		            firstline=Firstline+1
		          end if
		          
		          For i = minderror-2 To minderror+1 'save 4 profiles
		            if i<0 then
		              'skip
		            else
		              dom=u(i,isam)+ chr(9)+u(i,ik)
		              For j=ila to last+17
		                dom=dom+chr(9)+u(i,j)
		              next 
		              
		              'if  Single.FromString(u(i,ik))<1 then
		              'skip
		              Poutput.Writeline dom
		              'end if
		            end if
		          next
		          
		          Poutput.Close
		        Catch e As IOException
		          messagebox "Unable to append to file."
		        End Try
		        
		      end if
		      
		      minDerror=0
		      For i=nsstart+1 to ns 'clean out for next CoLa
		        s(i)=""
		        ki(i)=0
		        g(i) =0
		        for j=0 to last+26   '0  gets rid of model just made
		          u(i, j) = ""    'clear out previous model from u(i,j) matrix
		        next
		      next
		    end if
		    
		  next 'cola
		  
		  if imode=3 then 
		    Var Poutput As TextOutputStream
		    Try
		      Poutput = TextOutputStream.Open(tplusdump)
		      Poutput.WriteLine CDPTail
		      Poutput.Close
		    Catch e As IOException
		      messagebox "Unable to append to file."
		    End Try
		  end if
		  
		  if imode=2 or imode=3 then 
		    Var output As TextOutputStream
		    Try
		      output = TextOutputStream.Open(tdump)
		      output.WriteLine CDPTail
		      output.Close
		    Catch e As IOException
		      messagebox "Unable to append to file."
		    End Try
		    
		    minDerror=0
		    'clear out previous model from u(i,j) matrix etc
		    u.ResizeTo(-1,-1)
		    s.ResizeTo(-1)
		    ki.ResizeTo(-1)
		    g.ResizeTo(-1)
		    vn.ResizeTo(-1)
		    symscale.ResizeTo(-1)
		    
		    u.ResizeTo(irw , icl+12)
		    s.ResizeTo(irw +1000)
		    ki.ResizeTo(irw +1000)
		    g.ResizeTo(irw  + 1000)
		    vn.ResizeTo(irw +1000)
		    symscale.ResizeTo(irw +1000)
		    
		    ns=0 '
		    nfound=0
		    dom="Traverse has ended"+crlf
		    dom=dom + "To run a different model, reread the geochem data and replot the spiderdiagram."+crlf
		    dom=dom +"       "+crlf+prelimfails+crlf
		    messagebox dom
		  end if
		  
		  
		  yyb = 0.001
		  yye = 100
		  nspremod=ns   'HELP Nov11 brought this back to make Histograms work
		  zerosin = True
		  
		  'exception err
		  'If err IsA OutOfBoundsException Then
		  'MessageBox("The value of ns = "+str(ns)+"   nfound= "+str(nfound))
		  'End If
		End Sub
	#tag EndMethod


#tag EndWindowCode

#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		InitialValue=""
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimumHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximumHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Type"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Types"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasCloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasMinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasFullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Visible=false
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Visible=true
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="DefaultLocation"
		Visible=true
		Group="Behavior"
		InitialValue="2"
		Type="Locations"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Window Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackgroundColor"
		Visible=true
		Group="Background"
		InitialValue="&cFFFFFF"
		Type="ColorGroup"
		EditorType="ColorGroup"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		InitialValue=""
		Type="Picture"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		InitialValue=""
		Type="DesktopMenuBar"
		EditorType=""
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="False"
		Type="Boolean"
		EditorType=""
	#tag EndViewProperty
#tag EndViewBehavior
