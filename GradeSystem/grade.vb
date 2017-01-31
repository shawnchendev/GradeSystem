Public Class grade

    'class variable
    Private name As String
    Private id As Integer
    Private gpaindex As Integer
    Private satindex As Integer
    Private hsqindex As Integer
    Private docindex As Integer
    Private geo1 As Boolean
    Private geo2 As Boolean
    Private geo3 As Boolean
    Private Alun1 As Boolean
    Private Alun2 As Boolean
    Private essayindex As Integer
    Private LS1 As Boolean
    Private LS2 As Boolean
    Private LS3 As Boolean
    Private misindex As Integer

    ' class contructor
    Public Sub New()
        name = ""
        gpaindex = 0
        satindex = 0
        hsqindex = 0
        docindex = 0
        essayindex = 0
        misindex = 0
        geo1 = False
        geo2 = False
        geo3 = False
        Alun1 = False
        Alun2 = False
        LS1 = False
        LS2 = False
        LS3 = False
    End Sub
    ' set and get name 
    Public Property Name1 As String
        Get
            Return name
        End Get
        Set(ByVal value As String)
            name = value
        End Set
    End Property
    'set and get id
    Public Property ID1 As Integer
        Get
            Return id
        End Get
        Set(ByVal value As Integer)
            id = value
        End Set
    End Property
    'GPA 
    Public Property GPA_index As Integer
        Get
            Return gpaindex
        End Get
        Set(ByVal value As Integer)
            gpaindex = value
        End Set
    End Property

    Function getGPA() As Integer
        Return GPA_index * 2
    End Function

    'SAT
    Public Property SAT_index As Integer
        Get
            Return satindex
        End Get
        Set(ByVal value As Integer)
            satindex = value
        End Set
    End Property

    Function getSAT() As Integer
        Select Case SAT_index
            Case 0
                Return 0
            Case 1
                Return 6
            Case 2
                Return 10
            Case 3
                Return 11
            Case 4
                Return 12
            Case Else
                Return 0
        End Select
    End Function


    'High school quality  index
    Public Property HSQ_index As Integer
        Get
            Return hsqindex
        End Get
        Set(ByVal value As Integer)
            hsqindex = value
        End Set
    End Property
    ' return high school quanlity score
    Function getHSQ() As Integer
        Return HSQ_index * 2
    End Function

    'difficulty of curriculums index
    Public Property DOC_index As Integer
        Get
            Return docindex
        End Get
        Set(ByVal value As Integer)
            docindex = value
        End Set
    End Property
    ' renturn Difficulty of Curriculum score
    Function getDOC() As Integer
        Select Case DOC_index
            Case 0
                Return -4
            Case 1
                Return -2
            Case 2
                Return 0
            Case 3
                Return 2
            Case 4
                Return 4
            Case 5
                Return 6
            Case 6
                Return 8
            Case Else
                Return 0
        End Select
    End Function

    'geography index as boolean value
    Public Property Geo_geo1 As Boolean
        Get
            Return geo1
        End Get
        Set(ByVal value As Boolean)
            geo1 = value
        End Set
    End Property
    Public Property Geo_geo2 As Boolean
        Get
            Return geo2
        End Get
        Set(ByVal value As Boolean)
            geo2 = value
        End Set
    End Property
    Public Property Geo_geo3 As Boolean
        Get
            Return geo3
        End Get
        Set(ByVal value As Boolean)
            geo3 = value
        End Set
    End Property
    ' return Geography score
    Function getGeo() As Integer
        Dim score() As Integer = {0, 0, 0}
        If Geo_geo1 = True Then
            score(0) = 10
        End If
        If Geo_geo2 = True Then
            score(1) = 6
        End If
        If Geo_geo3 = True Then
            score(2) = 2
        End If
        Return score(0) + score(1) + score(2)
    End Function

    'alunmi index
    Public Property Alunmi_Alun1 As Boolean
        Get
            Return Alun1
        End Get
        Set(ByVal value As Boolean)
            Alun1 = value
        End Set
    End Property
    Public Property Alunmi_Alun2 As Boolean
        Get
            Return Alun2
        End Get
        Set(ByVal value As Boolean)
            Alun2 = value
        End Set
    End Property

    ' return Alunmi score
    Function getAlunmi() As Integer
        Dim alunscore() As Integer = {0, 0}
        If Alunmi_Alun1 = True Then
            alunscore(0) = 4

        End If
        If Alunmi_Alun2 = True Then
            alunscore(1) = 1
        End If
        Return alunscore(0) + alunscore(1)
    End Function
    'essay index
    Public Property Essay_index As Integer
        Get
            Return essayindex
        End Get
        Set(ByVal value As Integer)
            essayindex = value
        End Set
    End Property
    'renturn essay score
    Function getEssay() As Integer
        Return Essay_index + 1
    End Function

    'leadership and service index
    Public Property Leadership_LS1 As Boolean
        Get
            Return LS1
        End Get
        Set(ByVal value As Boolean)
            LS1 = value
        End Set
    End Property
    Public Property Leadership_LS2 As Boolean
        Get
            Return LS2
        End Get
        Set(ByVal value As Boolean)
            LS2 = value
        End Set
    End Property
    Public Property Leadership_LS3 As Boolean
        Get
            Return LS3
        End Get
        Set(ByVal value As Boolean)
            LS3 = value
        End Set
    End Property
    'retern Leadership and service score base of the selected index
    Function getLS() As Integer
        Dim lsscore() As Integer = {0, 0, 0}
        If Leadership_LS1 = True Then
            lsscore(0) = 1
        End If
        If Leadership_LS2 = True Then
            lsscore(1) = 2
        End If
        If Leadership_LS3 = True Then
            lsscore(2) = 5
        End If
        Return lsscore(0) + lsscore(1) + lsscore(2)
    End Function

    'miscellaneous selected index 
    Public Property mis_index As Integer
        Get
            Return misindex
        End Get
        Set(ByVal value As Integer)
            misindex = value
        End Set
    End Property
    ' return miscellaneous score base on the index
    Function getMIS() As Integer
        Select Case misindex
            Case 0
                Return 20
            Case 1
                Return 5
            Case 2
                Return 20
            Case 3
                Return 20
            Case Else
                Return 0
        End Select
    End Function

    'calculation total grade
    Function getTotal() As Integer
        Dim sectiontotal As Integer
        sectiontotal = getGeo() + getAlunmi() + getLS() + getEssay() + getMIS()
        If sectiontotal > 40 Then
            Return 40 + getGPA() + getSAT() + getHSQ() + getDOC()
        Else
            Return sectiontotal + getGPA() + getSAT() + getHSQ() + getDOC()
        End If
    End Function

End Class