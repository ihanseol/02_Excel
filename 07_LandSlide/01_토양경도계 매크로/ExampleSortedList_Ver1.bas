Attribute VB_Name = "ExampleSortedList_Ver1"
Option Explicit

'http://egloos.zum.com/timebird/v/7403799
'http://timebird.egloos.com/category/%EC%98%A4%ED%94%BC%EC%8A%A4%2FVBA%2FOffice.JS

Private Sub demoArrayList()
    Dim arrList As Object
    Dim item
    
    'Create the ArrayList
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    arrList.Add "Hello"
    arrList.Add "You"
    arrList.Add "There"
    arrList.Add "Man"
    arrList.Remove "Man"
    
    'Get number of items
    Debug.Print arrList.Count 'Result: 3
    
    For Each item In arrList
      Debug.Print item
    Next
    
End Sub

Private Sub demoSortedList()
    Dim sortedList As Object
    
    ' Create the SortedList
    Set sortedList = CreateObject("System.Collections.SortedList")
    
    sortedList.Add "ThisortedListrd", "!"
    sortedList.Add "Second", "World"
    sortedList.Add "First", "Hello"

    ' Displays the properties and values of the SortedList.
    Debug.Print "Count:"; sortedList.Count
    Debug.Print "Capacity:"; sortedList.Capacity
    
    Dim i As Long
    
    For i = 0 To sortedList.Count - 1
        Debug.Print sortedList.GetKey(i), sortedList.GetByIndex(i)
    Next
End Sub

Private Sub demoQueue()
    Dim queue As Object
    Dim peekAtFirst, doesContain, firstInQueue
    
    'Create the Queue
    Set queue = CreateObject("System.Collections.Queue")
    
    queue.Enqueue "Hello"
    queue.Enqueue "There"
    queue.Enqueue "Mr"
    queue.Enqueue "Smith"
    
    peekAtFirst = queue.Peek() 'Result" "Hello"
    Debug.Print peekAtFirst
    
    doesContain = queue.Contains("htrh") 'Result: False
    Debug.Print doesContain
    
    doesContain = queue.Contains("Hello") 'Result: True
    Debug.Print doesContain
    
    'Get first item in Queue and remove it from the Queue
    firstInQueue = queue.Dequeue() '"Hello"
    Debug.Print firstInQueue
    
    'Count items
    Debug.Print queue.Count 'Result: 3
    
    'Clear the Queue
    queue.Clear
    
    Set queue = Nothing
End Sub

Private Sub demoStack()
    Dim stack As Object
    Dim peekAtTopOfStack, doesContain, topStack
    
    'Create Stack
    Set stack = CreateObject("System.Collections.Stack")
    
    stack.Push "Hello"
    stack.Push "There"
    stack.Push "Mr"
    stack.Push "Smith"
    
    peekAtTopOfStack = stack.Peek()
    Debug.Print peekAtTopOfStack
    
    doesContain = stack.Contains("htrh") 'Result: False
    Debug.Print doesContain
    
    doesContain = stack.Contains("Hello") 'Result: True
    Debug.Print doesContain
    
    'Get item from the top of the stack (LIFO)
    topStack = stack.Pop()  'Result: "Smith"
    Debug.Print topStack
    
    'Clear the Stack
    stack.Clear
    
    Set stack = Nothing
End Sub
