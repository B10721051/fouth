Attribute VB_Name = "Module1"
Option Explicit

Public Function eoq(demand, holding, fixed) As Integer
eoq = ((2 * demand * fixed) / holding) ^ (1 / 2)
End Function
