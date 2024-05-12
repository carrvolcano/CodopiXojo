#tag DesktopWindow
Begin DesktopWindow SIAdjustWindow
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
   MenuBar         =   407050239
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

		Sub SIAdjust(meltmode() as double, mantmode() as double, source() as double)
		  var i,j as integer
		  var SIpm(15),SIdm(15),f(15),ch(16),maxmelt,dum as double
		  var c(15,15) as double
		  var xsum, ysum, xysum, xsqsum, ysqsum, xdb, ydb As Double
		  var slope(2,20), cept(2,20), sloperr, cepterr, xstderr,xv(15),yv(15) As Double
		  var n,m as integer
		  
		  'Subroutine to calculate slope amd intercept adjustments needed by assumption that the Part.Coeff. of La=0
		  ch(1)=0.237 'La
		  ch(2)= 0.613  'Ce
		  ch(3)=0.0928  'Pr
		  ch(4)=0.457  'Nd
		  ch(5)=0.148  'Sm
		  ch(6)=0.0563  'Eu
		  ch(7)=0.199  'Gd
		  ch(8)=0.0361  'Tb
		  ch(9)=0.246  'Dy
		  ch(10)=0.0546  'Ho
		  ch(11)=0.16  'Er
		  ch(12)=0.0247  'Tm
		  ch(13)=0.161  'Yb
		  ch(14)=0.0246 'Lu
		  
		  'max melt - needed to set range of melt %s
		  maxmelt=100
		  for i=1 to 4   '4 minerals
		    if meltmode(i)>0 and mantmode(i)>0 then
		      dum=mantmode(i)/meltmode(i)
		      if dum<maxmelt then
		        maxmelt=dum
		      end if
		    end if
		  next
		  
		  if maxmelt>.2 then
		    maxmelt=.2
		  end if
		  
		  'calculate  d's + p's
		  For i = 1 To nme   'for each element in pc file
		    SIdm(i) = 0
		    SIpm(i) = 0
		    For j = 1 To nmn
		      SIdm(i) = SIdm(i) + ptcf(i, j) * mantmode(j)/100
		      SIpm(i) = SIpm(i) + ptcf(i, j) * meltmode(j)/100
		    Next 
		  next
		  SIdm(0)=0' La PC's=0
		  SIpm(0)=0' La PC's=0
		  
		  source(0)=source(1)' will use source(0) as La for when La PC's=0
		  
		  for i=1 to 8
		    f(i)=i*maxmelt/8
		  next
		  
		  for i= 1 to 8  'melts  a fixed #
		    for j= 0 to nme 'elements
		      c(i,j)=source(j)/(SIdm(j) + f(i) * (1 - SIpm(j)))' batch melt eqn using estimated mantle mode and melt mode of trial
		    next
		    'note for La PC's=0,  c(i,0)=source(j)/ f(i) 
		  next
		  
		  'Start Regressions
		  'Normal version
		  'dom="Normal version La ch=  "+str(ch(1))+crlf
		  
		  for m= 2 to nme 'i elements Ce to Lu
		    for j=1 to 8 'j melts
		      xv(j)=c(j,1)/ch(1)
		      yv(j)=xv(j)/(c(j,m)/ch(m))
		    next
		    
		    n=8
		    xsum = 0
		    ysum = 0
		    xysum = 0
		    xsqsum = 0
		    ysqsum = 0
		    For i = 1 To 8  'melts
		      xdb = xv(i)
		      ydb = yv(i)
		      xsum = xsum + xdb
		      ysum = ysum + ydb
		      xysum = xysum + xdb * ydb
		      xsqsum = xsqsum + xdb * xdb
		      ysqsum = ysqsum + ydb * ydb
		    Next
		    
		    If nme < 3 Then
		      MessageBox "Need at least 3 points"
		      Return 
		    End If
		    if xsum.Equals(0, 1) then
		      
		    else
		      slope(1,m) = (n * xysum - xsum * ysum) / (n * xsqsum - xsum * xsum)
		      cept(1,m) = (ysum / n) - slope(1,m) * (xsum / n)
		      xstderr = (ysqsum - cept(1,m) * ysum - slope(1,m) * xysum) / (n - 2)
		      If xstderr <= 0 Then
		        xstderr = 0
		      Else
		        xstderr = pow(xstderr,.5)
		      End If
		      cepterr = xstderr * pow((1 / n + (xsum * xsum / n) / (n * xsqsum - xsum * xsum)),.5)
		      sloperr = xstderr / pow(((n * xsqsum - xsum * xsum) / n),.5)
		      
		    end if
		  next m
		  
		  'LaZero version
		  'dom=" Slopes and Ints La zero"
		  for m= 2 to nme 'i elements Ce to Lu
		    for j=1 to 8 'j melts
		      xv(j)=c(j,0)/ch(1)  'La
		      yv(j)=xv(j)/(c(j,m)/ch(m))
		    next
		    
		    n=8
		    xsum = 0
		    ysum = 0
		    xysum = 0
		    xsqsum = 0
		    ysqsum = 0
		    For i = 1 To 8  'melts
		      xdb = xv(i)
		      ydb = yv(i)
		      xsum = xsum + xdb
		      ysum = ysum + ydb
		      xysum = xysum + xdb * ydb
		      xsqsum = xsqsum + xdb * xdb
		      ysqsum = ysqsum + ydb * ydb
		    Next
		    
		    If nme < 3 Then
		      MessageBox "Need at least 3 points"
		      Return 
		    End If
		    
		    if xsum.Equals(0, 1) then
		      
		    else
		      
		      slope(2,m) = (n * xysum - xsum * ysum) / (n * xsqsum - xsum * xsum)
		      cept(2,m) = (ysum / n) - slope(2,m) * (xsum / n)
		      xstderr = (ysqsum - cept(2,m) * ysum - slope(2,m) * xysum) / (n - 2)
		      If xstderr <= 0 Then
		        xstderr = 0
		      Else
		        xstderr = pow(xstderr,.5)
		      End If
		      cepterr = xstderr * pow((1 / n + (xsum * xsum / n) / (n * xsqsum - xsum * xsum)),.5)
		      sloperr = xstderr / pow(((n * xsqsum - xsum * xsum) / n),.5)
		      
		    end if 
		  next m
		  
		  for m=2 to nme 'ignore La
		    SAdjust(m-1)=slope(2,m)/slope(1,m)
		  next
		  
		  SAdjust(nme+1)=cept(2,3)/cept(1,3)
		  if cept(1,3).Equals(0, 1) then
		    'skip
		  else
		    IntAdj=cept(2,3)/cept(1,3)
		  end if
		  
		  return 
		  
		End Sub
	#tag EndMethod


#tag EndWindowCode

