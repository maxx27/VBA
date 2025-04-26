Attribute VB_Name = "Коллекции"
Option Explicit


' Проверить вхождения элемента в коллекцию.

Function ВходитВКоллекцию(Элемент As Variant, Коллекция As Collection) As Boolean
    Dim ОчереднойЭлемент As Variant

    For Each ОчереднойЭлемент In Коллекция
        If ОчереднойЭлемент = Элемент Then
            ВходитВКоллекцию = True
            Exit Function
        End If
    Next
    
    ВходитВКоллекцию = False
End Function

' Получить только уникальные элементы коллекции.

Function УникальныеИзКоллекции(Коллекция As Collection) As Collection
    Dim КоллекцияУникальных As New Collection
    Dim Элемент As Variant
    
    If Коллекция Is Nothing Then
        Set УникальныеИзКоллекции = Nothing
        Exit Function
    End If
    
    For Each Элемент In Коллекция
        If Not ВходитВКоллекцию(Элемент, КоллекцияУникальных) Then КоллекцияУникальных.Add (Элемент)
    Next
    
    Set УникальныеИзКоллекции = КоллекцияУникальных
End Function

' Сортирует элементы коллекции методом пузырька
' TODO: стоит добавить сравнение типов элементов (тогда предварительно сортировать по типу, а затем внутри типа)

Function СортироватьКоллекцию(Коллекция As Collection) As Collection
    Dim БылаПерестановка As Boolean
    Dim Временный As Variant
    Dim i As Integer
    
    If Коллекция Is Nothing Then
        Set СортироватьКоллекцию = Nothing
        Exit Function
    End If
    
    Do
        БылаПерестановка = False
        For i = 2 To Коллекция.Count
            If Коллекция.Item(i - 1) > Коллекция.Item(i) Then
                Коллекция.Add Коллекция.Item(i), Before:=i - 1
                Коллекция.Remove i + 1
                БылаПерестановка = True
            End If
        Next
    Loop While БылаПерестановка
    
    Set СортироватьКоллекцию = Коллекция
End Function


