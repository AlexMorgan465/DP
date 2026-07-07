Sub GenererPlanningAFEDIM(collabs() As Collaborateur, nb As Integer)
    SchedulingEngine.GenerateProjectPlanning ProjectRulesRepository.LoadProjectRule("AFEDIM"), collabs, nb
End Sub
Sub GenererPlanningACCESSIBILITE(collabs() As Collaborateur, nb As Integer)
    SchedulingEngine.GenerateProjectPlanning ProjectRulesRepository.LoadProjectRule("ACCESSIBILITE"), collabs, nb
End Sub
Sub GenererPlanningCMLEASING(collabs() As Collaborateur, nb As Integer)
    SchedulingEngine.GenerateProjectPlanning ProjectRulesRepository.LoadProjectRule("CMLEASING"), collabs, nb
End Sub
Sub GenererPlanningGLF(collabs() As Collaborateur, nb As Integer)
    SchedulingEngine.GenerateProjectPlanning ProjectRulesRepository.LoadProjectRule("GLF"), collabs, nb
End Sub
