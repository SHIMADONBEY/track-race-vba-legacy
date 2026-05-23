Attribute VB_Name = "CustomError"
'namespace=vba-files/track-race/util
Option Explicit
Option Private Module

Public Enum CustomErrorCodeEnum
    UndefinedError = vbObjectError + 256 * 2
    InvalidRoundCount
    DuplicateRoundName
    QualifiersNotAssigned
    ResultAlreadyRegistred
    AllGroupNotConfirmed
    NoQualifiers
    NoEntryAssigned
    EntryOrderDuplicated
    ResultOrderDuplicated
    TargetRoundNotSelected
End Enum
