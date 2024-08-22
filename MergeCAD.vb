
filePath1="C:\1.dwg"
filePath2="C:\2.dwg"
Set acadHost = CreateObject("TeighaX.OdaHostApp")
Set aryDouble=createOBject("Custom.CreateDoubleArray.DoubleArray")
Set app = acadHost.Application
Set acadDoc1=app.documents.open(filePath1)
Set database1=acadDoc1.database
Set acadDoc2=app.documents.open(filePath2)
Set database2=acadDoc2.database 
offx=600
offy=0
docIndex=1 '合并文件的数量，以便计算复制后放置的坐标
Set  TeighaServiceHelper=createOBject("Custom.TeighaService.Helper")'编写一个dll用于将vbs 动态数组转换为double数组
'执行合并函数，将文件2合并到文件1中，多份文件可以依次类推
call MergeDocument(acadDoc1,acadDoc2,TeighaServiceHelper,offx,offy,docIndex)
acadDoc1.Save
cadDoc1.close
app.quit
Set app=Nothing
msgbox "完成",64,"提示" 

Function MergeDocument(acadDoc1,acadDoc2,TeighaServiceHelper,offx,offy,docIndex)
   Set database1=acadDoc1.database
   Set database2=acadDoc2.database
   Dim fromPoint(2)
    Dim ToPoint(2)
    fromPoint(0)=0 
    fromPoint(1)=0
    fromPoint(2)=0
    ToPoint(0)=offx '计算合并后的位置
    ToPoint(1)=offy
    ToPoint(2)=0
	'因vbs不支持double数组定义，需要将数组转换为double数组
    fromPoint1=TeighaServiceHelper.Convert2DoubleArray(fromPoint) 
    ToPoint1=TeighaServiceHelper.Convert2DoubleArray(ToPoint)
    '先把图纸2中块的名字批量改掉，以免合并后因重名导致引用错误，有时不同文档相同块名但内容不同。
    For i=2 To database2.blocks.Count-1 
        Set obj=database2.blocks.Item(i)
         If instr(ucase(obj.Name),"SPACE")<=0 Then '跳过modelspace和paperspace
          obj.Name=docIndex&"_"&obj.Name      
        For k=0 To obj.Count-1
          If obj(k).EntityName  ="AcDbBlockReference" Then
              strName=obj(k).EffectiveName
              If instr(strName,docIndex&"_")<=0 Then '批量改名
                obj(k).Name=docIndex&"_"&strName
              End If
          End If
        Next     
      End If
    Next
	'复制图纸2中的内容到图纸1
    For i=0 To acadDoc2.modelspace.Count-1
        Set obj=acadDoc2.modelspace.Item(i)
        ary1= database2.CopyObjects( obj,acadDoc1.modelspace) 
        For j=0 To ubound(ary1)
          Set objSub=ary1(j)
          objSub.move fromPoint1,ToPoint1

        Next
        
    Next
 
   acadDoc2.close
End Function


