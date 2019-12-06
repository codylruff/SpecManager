Attribute VB_Name = "Enums"
Option Explicit

' LOG TYPES
Public Enum LogType
    RuntimeLog = 0
    UserLog = 1
    ErrorLog = 3
    TestLog = 4
    DebugLog = 5
    SqlLog = 6
    ExportLog = 7
    RevisionLog = 8
End Enum

' LOG VERBOSITY LEVELS
Public Enum LogLevel
    LOG_LOW = 3
    LOG_TEST = 4
    LOG_DEBUG = 6
    LOG_ALL = 8
End Enum

Public Enum InputBoxType
    Password = 1 ' Masked using systempassword mask
    SingleLineText = 2 ' Single line text
    MultiLineText = 32 ' Multi line text
    Number = 4 ' Numbers only
    ShortDate = 4 ' Masked dd/mm/yyyy. Dates are validated upon exit
    LongDate = 16 ' asked using dd/Month/yyyy
    DateTime = 48 ' masked using dd/mm/yyyy hh:mm:ss
End Enum

Public Enum DocumentPackageVariant
    Default = 0
    WeavingStyleChange = 1
    WeavingTieBack = 2
    FinishingWithQC = 3
    FinishingNoQC = 4
    Isotex = 5
End Enum

Public Enum PixelDirection
    Horizontal
    Vertical
End Enum

Public Enum symbology
        '/// <summary>
        '/// Code One 2D symbol.
        '/// </summary>
        CodeOne = 0
        '/// <summary>
        '/// Code 39 (ISO 16388)
        '/// </summary>
        Code39 = 1
        '/// <summary>
        '/// Code 39 extended ASCII.
        '/// </summary>
        Code39Extended = 2
        '/// <summary>
        '/// Logistics Applications of Automated Marking and Reading Symbol.
        '/// </summary>
        LOGMARS = 3
        '/// <summary>
        '/// Code 32 (Italian Pharmacode)
        '/// </summary>
        Code32 = 4
        '/// <summary>
        '/// Pharmazentralnummer (PZN - German Pharmaceutical Code)
        '/// </summary>
        PharmaZentralNummer = 5
        '/// <summary>
        '/// Pharmaceutical Binary Code.
        '/// </summary>
        Pharmacode = 6
        '/// <summary>
        '/// Pharmaceutical Binary Code (2 Track)
        '/// </summary>
        Pharmacode2Track = 7
        '/// <summary>
        '/// Code 93
        '/// </summary>
        Code93 = 8
        '/// <summary>
        '/// Channel Code.
        '/// </summary>
        ChannelCode = 9
        '/// <summary>
        '/// Telepen Code.
        '/// </summary>
        Telepen = 10
        '/// <summary>
        '/// Telepen Numeric Code.
        '/// </summary>
        TelepenNumeric = 11
        '/// <summary>
        '/// Code 128/GS1-128 (ISO 15417)
        '/// </summary>
        Code128 = 12
        '/// <summary>
        '/// European Article Number (14)
        '/// </summary>
        EAN14 = 13
        '/// <summary>
        '/// Serial Shipping Container Code.
        '/// </summary>
        SSCC18 = 14
        '/// <summary>
        '/// Standard 2 of 5 Code.
        '/// </summary>
        Standard2of5 = 15
        '/// <summary>
        '/// Interleaved 2 of 5 Code.
        '/// </summary>
        Interleaved2of5 = 16
        '/// <summary>
        '/// Matrix 2 of 5 Code.
        '/// </summary>
        Matrix2of5 = 17
        '/// <summary>
        '/// IATA 2 of 5 Code.
        '/// </summary>
        IATA2of5 = 18
        '/// <summary>
        '/// Datalogic 2 of 5 Code.
        '/// </summary>
        DataLogic2of5 = 19
        '/// <summary>
        '/// ITF 14 (GS1 2 of 5 Code)
        '/// </summary>
        ITF14 = 20
        '/// <summary>
        '/// Deutsche Post Identcode (DHL)
        '/// </summary>
        DeutschePostIdentCode = 21
        '/// <summary>
        '/// Deutsche Post Leitcode (DHL)
        '/// </summary>
        DeutshePostLeitCode = 22
        '/// <summary>
        '/// Codabar Code.
        '/// </summary>
        Codabar = 23
        '/// <summary>
        '/// MSI Plessey Code.
        '/// </summary>
        MSIPlessey = 24
        '/// <summary>
        '/// UK Plessey Code.
        '/// </summary>
        UKPlessey = 25
        '/// <summary>
        '/// Code 11.
        '/// </summary>
        Code11 = 26
        '/// <summary>
        '/// International Standard Book Number.
        '/// </summary>
        ISBN = 27
        '/// <summary>
        '/// European Article Number (13)
        '/// </summary>
        EAN13 = 28
        '/// <summary>
        '/// European Article Number (8)
        '/// </summary>
        EAN8 = 29
        '/// <summary>
        '/// Universal Product Code (A)
        '/// </summary>
        UPCA = 30
        '/// <summary>
        '/// Universal Product Code (E)
        '/// </summary>
        UPCE = 31
        '/// <summary>
        '/// GS1 Databar Omnidirectional.
        '/// </summary>
        DatabarOmni = 32
        '/// <summary>
        '/// GS1 Databar Omnidirectional Stacked.
        '/// </summary>
        DatabarOmniStacked = 33
        '/// <summary>
        '/// GS1 Databar Stacked.
        '/// </summary>
        DatabarStacked = 34
        '/// <summary>
        '/// GS1 Databar Omnidirectional Truncated.
        '/// </summary>
        DatabarTruncated = 35
        '/// <summary>
        '/// GS1 Databar Limited.
        '/// </summary>
        DatabarLimited = 36
        '/// <summary>
        '/// GS1 Databar Expanded.
        '/// </summary>
        DatabarExpanded = 37
        '/// <summary>
        '/// GS1 Databar Expanded Stacked.
        '/// </summary>
        DatabarExpandedStacked = 38
        '/// <summary>
        '/// Data Matrix (ISO 16022)
        '/// </summary>
        DataMatrix = 39
        '/// <summary>
        '/// QR Code (ISO 18004)
        '/// </summary>
        QRCode = 40
        '/// <summary>
        '/// Micro variation of QR Code.
        '/// </summary>
        MicroQRCode = 41
        '/// <summary>
        '/// UPN variation of QR Code.
        '/// </summary>
        UPNQR = 42
        '/// <summary>
        '/// Aztec (ISO 24778)
        '/// </summary>
        Aztec = 43
        '/// <summary>
        '/// Aztec Runes.
        '/// </summary>
        AztecRunes = 44
        '/// <summary>
        '/// Maxicode (ISO 16023)
        '/// </summary>
        MaxiCode = 45
        '/// <summary>
        '/// PDF417 (ISO 15438)
        '/// </summary>
        PDF417 = 46
        '/// <summary>
        '/// PDF417 Truncated.
        '/// </summary>
        PDF417Truncated = 47
        '/// <summary>
        '/// Micro PDF417 (ISO 24728)
        '/// </summary>
        MicroPDF417 = 48
        '/// <summary>
        '/// Australia Post Standard.
        '/// </summary>
        AusPostStandard = 49
        '/// <summary>
        '/// Australia Post Reply Paid.
        '/// </summary>
        AusPostReplyPaid = 50
        '/// <summary>
        '/// Australia Post Redirect.
        '/// </summary>
        AusPostRedirect = 51
        '/// <summary>
        '/// Australia Post Routing.
        '/// </summary>
        AusPostRouting = 52
        '/// <summary>
        '/// United States Postal Service Intelligent Mail.
        '/// </summary>
        USPS = 53
        '/// <summary>
        '/// PostNET (Postal Numeric Encoding Technique)
        '/// </summary>
        PostNet = 54
        '/// <summary>
        '/// Planet (Postal Alpha Numeric Encoding Technique)
        '/// </summary>
        Planet = 55
        '/// <summary>
        '/// Korean Post.
        '/// </summary>
        KoreaPost = 56
        '/// <summary>
        '/// Facing Identification Mark (FIM)
        '/// </summary>
        FIM = 57
        '/// <summary>
        '/// UK Royal Mail 4 State Code.
        '/// </summary>
        RoyalMail = 58
        '/// <summary>
        '/// KIX Dutch 4 State Code.
        '/// </summary>
        KixCode = 59
        '/// <summary>
        '/// DAFT Code (Generic 4 State Code)
        '/// </summary>
        DaftCode = 60
        '/// <summary>
        '/// Flattermarken (Markup Code)
        '/// </summary>
        Flattermarken = 61
        '/// <summary>
        '/// Japanese Post.
        '/// </summary>
        JapanPost = 62
        '/// <summary>
        '/// Codablock-F 2D symbol.
        '/// </summary>
        CodablockF = 63
        '/// <summary>
        '/// Code 16K 2D symbol.
        '/// </summary>
        Code16K = 64
        '/// <summary>
        '/// Dot Code 2D symbol.
        '/// </summary>
        DotCode = 65
        '/// <summary>
        '/// Grid Matrix 2D symbol.
        '/// </summary>
        GridMatrix = 66
        '/// <summary>
        '/// Code 49 2D symbol.
        '/// </summary>
        Code49 = 67
        '/// <summary>
        '/// Han Xin 2D symbol.
        '/// </summary>
        HanXin = 68

        '/// <summary>
        '/// VIN code symbol.
        '/// </summary>
        VINCode = 69

        '/// <summary>
        '/// Mailmark 4 state postal.
        '/// </summary>
        RoyalMailMailmark = 70

        '/// <summary>
        '/// Not a valid Symbol ID.
        '/// </summary>
        Invalid = -1
End Enum

'FTP enums
Public Enum ProtocolEnum
    Sftp = 0
    Scp = 1
    ftp = 2
    Webdav = 3
    S3 = 4
End Enum

Public Enum FtpSecureEnum
    None = 0
    Implicit = 1
    Explicit = 3
End Enum

'end of ftp enums

Public Enum JsonFormatting
    '
    ' Summary:
    '     Specifies formatting options for the Newtonsoft.Json.JsonTextWriter.
    '
    ' Summary:
    '     No special formatting is applied. This is the default.
    None = 0
    '
    ' Summary:
    '     Causes child objects to be indented according to the Newtonsoft.Json.JsonTextWriter.Indentation
    '     and Newtonsoft.Json.JsonTextWriter.IndentChar settings.
    Indented = 1
End Enum


Public Enum AcImportXMLOption
    'Creates a new table based on the structure of the specified XML file.
    acStructureOnly = 0
    'Imports the data into a new table based on the structure of the specified XML file.
    acStructureAndData = 1
    
    acAppendData = 2      'Imports the data into an existing table.
End Enum
