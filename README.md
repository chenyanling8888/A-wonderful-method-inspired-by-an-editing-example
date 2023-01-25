# A-wonderful-method-inspired-by-an-editing-example
作者:...、凡超、......

今天,发现编译过程,清晰度高,高得惊人,还有一个是数理考虑周全,能拆亦能合,即,人人能做到自主随意的要拆分使用就拆分使用,要复合使用就使用,重要的是其原理简直是奇妙不可言.奇妙不可言的同时还带给我很大的惊喜_收获一启发,就类似中国博大精深的古老文化的道理,简直神奇!收获这一启发理顺一个项目全局的子方法外还附带一个选项吧!

比如:

电子表格中边框这一选项中有:下框线,上框线,左框线,右框线,无框线,所有框线,外侧框线,粗底框线,双底框线,上下框线,......它们即能拆分使用,也能复合使用,神奇!

请看示例1:

Sub Miraculous_Border1()

    Range("A1:K74").select
    
    Selection.Borders(xlDiagonalDown).LineSytle=xlNone
    
    Selection.Borders(xlDiagonalUp).LineStytle=xlNone
    
    With selection.Border(xlEdgeLeft)
    
        .LineStyle=xlContinuous
        
        .ColorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle=Xlcontinuous
        
        .ColorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
    End With
    
    With Selection.Borders(xlEdgeRight)
    
        .LineStyle=xlcontinuous
        
        .ColorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
    End With
    
    With Selection.Borders(xlEdgeBottom)
    
        .LineStyle=xlcontinuous
        
        .ColorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
    End With
    
    With Selection.Borders(xlInsideVertical)
    
        .LineStyle=xlContinuous
        
        .ColorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        
        .LineStyle=xlContinuous
        
        .CorlorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
   End With
End Sub
   
示例1代码的运行结果是一个所有框线里的其中一种类型的电子表,那么,请在看看运行结果只有上下框线的代码,便能了知其中清晰度.至于那神奇的妙法,请容我有时间一一细数啊!

请看示例2:

Sub Miraculous_Border2()
    Range("B2:L2")
    
    With Selection.Borders(xlEdgeTop)
    
        .LineStyle=xlContinuous
        
        .CorlorIndex=0
        
        .TintAndShade=0
        
        .Weight=xlThin
        
    End With
    
    With Selection.Borders(xlEdgeBottom)
    
        .LineSty=xlContinuous
        
        .CorlorIndex=0
        
        .TintAndShade=0
        
        Weight=xlThin
        
   End With
End Sub

示例2代码的运行结果是一个只有上下框线的代码,便能了知其中清晰度,至于那神奇的妙法,请容我有时间一一细数!
事实上写这下这一段是帮助我记忆及引用的,因为学习编程是为能做项目,这一启发让我更接近目的地了.很开心这个发现!


    





    
    



