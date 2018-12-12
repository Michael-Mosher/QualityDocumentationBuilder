Option Explicit

 

Dim proc As Double

Dim this_esat As Double

Dim eval_date As Date

Dim eval_type As String

Dim verification As Boolean

Dim current_agent As String

Dim secondValid As Boolean

 

Public Sub Class_Initialize()

    verification = True

  SecondaryValid = True

End Sub

 

Public Property Let procedural(score As Double)

    proc = score

End Property

 

Public Property Let esat(score As Double)

    this_esat = score

End Property

 

Public Property Let edate(edate As Date)

    eval_date = edate

End Property

 

Public Property Let etype(etype As String)

    eval_type = etype

End Property

 

Public Property Let everification(everification As Boolean)

    verification = everification

End Property

 

Public Property Get etype() As String

    etype = eval_type

End Property

 

Public Property Get edate() As Date

    edate = eval_date

End Property

 

Public Property Get esat() As Double

    esat = this_esat

End Property

 

Public Property Get procedural() As Double

    procedural = proc

End Property

 

Public Property Get everification() As Boolean

    everification = verification

End Property

 

Public Property Let Agent(sAgentName As String)

    current_agent = sAgentName

End Property

 

Public Property Get Agent() As String

    Agent = current_agent

End Property

 

Public Property Let SecondaryValid(validity As Boolean)

  secondValid = validity

End Property

 

Public Property Get SecondaryValid() As Boolean

  SecondaryValid = secondValid

End Property
