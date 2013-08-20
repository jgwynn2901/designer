QA:
FNSConnectString_QA.reg
FNSUseAltDestination_True.reg



SPR0S2T and QA:
FNSEditCompletedCalls_True.reg
FNSSystemMonitor_True.reg




Special Case:
FNSSQLTrappMode_True.reg


Spell Check
Standard call center server
Interactive on, Passive on.  
Call reps may interactively check the spelling within a call.  They can change words, by selecting from the online list or entering a new word during the check.  They may also ignore misspellings.  When the call is closed, spell checking is performed again, unseen by the call rep.  If there are any words in a field marked for checking that are not in the dictionary, the call is marked as SPELL_ERR.

Standard spell checker server
Interactive on, Passive OFF. 
Spell checkers edit calls marked as SPELL_ERR.  They can change words, by selecting from the online list or entering a new word during the check.  They may also ignore misspellings.  When the call is closed, spell checking is skipped, unseen by the call rep.  The call is marked as COMPLETE.
