Attribute VB_Name = "modTimer"
Option Explicit


   Declare Sub GetLocalTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME)


   Type SYSTEMTIME
       wYear As Integer
       wMonth As Integer
       wDayOfWeek As Integer
       wDay As Integer
       wHour As Integer
       wMinute As Integer
       wSecond As Integer
       wMilliseconds As Integer
       End Type
                                                   
                

   Public Sub Sleep(Seconds As Integer)
   


   Dim Start As Long
       Start = Timer


       Do While Timer < Start + Seconds


           DoEvents ' Yield to other processes.
           Loop



  
   End Sub


