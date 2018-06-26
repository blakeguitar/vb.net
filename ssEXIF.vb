
'test

Public Class ssEXIF


    Private _Image As System.Drawing.Bitmap
    Private _Encoding As System.Text.Encoding = System.Text.Encoding.UTF8

    Public Enum TagNames As Integer
        ExifIFD = &H8769
        GpsIFD = &H8825
        NewSubfileType = &HFE
        SubfileType = &HFF
        ImageWidth = &H100
        ImageHeight = &H101
        BitsPerSample = &H102
        Compression = &H103
        PhotometricInterp = &H106
        ThreshHolding = &H107
        CellWidth = &H108
        CellHeight = &H109
        FillOrder = &H10A
        DocumentName = &H10D
        ImageDescription = &H10E
        EquipMake = &H10F
        EquipModel = &H110
        StripOffsets = &H111
        Orientation = &H112
        SamplesPerPixel = &H115
        RowsPerStrip = &H116
        StripBytesCount = &H117
        MinSampleValue = &H118
        MaxSampleValue = &H119
        XResolution = &H11A
        YResolution = &H11B
        PlanarConfig = &H11C
        PageName = &H11D
        XPosition = &H11E
        YPosition = &H11F
        FreeOffset = &H120
        FreeByteCounts = &H121
        GrayResponseUnit = &H122
        GrayResponseCurve = &H123
        T4Option = &H124
        T6Option = &H125
        ResolutionUnit = &H128
        PageNumber = &H129
        TransferFuncition = &H12D
        SoftwareUsed = &H131
        DateTime = &H132
        Artist = &H13B
        HostComputer = &H13C
        Predictor = &H13D
        WhitePoint = &H13E
        PrimaryChromaticities = &H13F
        ColorMap = &H140
        HalftoneHints = &H141
        TileWidth = &H142
        TileLength = &H143
        TileOffset = &H144
        TileByteCounts = &H145
        InkSet = &H14C
        InkNames = &H14D
        NumberOfInks = &H14E
        DotRange = &H150
        TargetPrinter = &H151
        ExtraSamples = &H152
        SampleFormat = &H153
        SMinSampleValue = &H154
        SMaxSampleValue = &H155
        TransferRange = &H156
        JPEGProc = &H200
        JPEGInterFormat = &H201
        JPEGInterLength = &H202
        JPEGRestartInterval = &H203
        JPEGLosslessPredictors = &H205
        JPEGPointTransforms = &H206
        JPEGQTables = &H207
        JPEGDCTables = &H208
        JPEGACTables = &H209
        YCbCrCoefficients = &H211
        YCbCrSubsampling = &H212
        YCbCrPositioning = &H213
        REFBlackWhite = &H214
        ICCProfile = &H8773
        Gamma = &H301
        ICCProfileDescriptor = &H302
        SRGBRenderingIntent = &H303
        ImageTitle = &H320
        Copyright = &H8298
        ResolutionXUnit = &H5001
        ResolutionYUnit = &H5002
        ResolutionXLengthUnit = &H5003
        ResolutionYLengthUnit = &H5004
        PrintFlags = &H5005
        PrintFlagsVersion = &H5006
        PrintFlagsCrop = &H5007
        PrintFlagsBleedWidth = &H5008
        PrintFlagsBleedWidthScale = &H5009
        HalftoneLPI = &H500A
        HalftoneLPIUnit = &H500B
        HalftoneDegree = &H500C
        HalftoneShape = &H500D
        HalftoneMisc = &H500E
        HalftoneScreen = &H500F
        JPEGQuality = &H5010
        GridSize = &H5011
        ThumbnailFormat = &H5012
        ThumbnailWidth = &H5013
        ThumbnailHeight = &H5014
        ThumbnailColorDepth = &H5015
        ThumbnailPlanes = &H5016
        ThumbnailRawBytes = &H5017
        ThumbnailSize = &H5018
        ThumbnailCompressedSize = &H5019
        ColorTransferFunction = &H501A
        ThumbnailData = &H501B
        ThumbnailImageWidth = &H5020
        ThumbnailImageHeight = &H502
        ThumbnailBitsPerSample = &H5022
        ThumbnailCompression = &H5023
        ThumbnailPhotometricInterp = &H5024
        ThumbnailImageDescription = &H5025
        ThumbnailEquipMake = &H5026
        ThumbnailEquipModel = &H5027
        ThumbnailStripOffsets = &H5028
        ThumbnailOrientation = &H5029
        ThumbnailSamplesPerPixel = &H502A
        ThumbnailRowsPerStrip = &H502B
        ThumbnailStripBytesCount = &H502C
        ThumbnailResolutionX = &H502D
        ThumbnailResolutionY = &H502E
        ThumbnailPlanarConfig = &H502F
        ThumbnailResolutionUnit = &H5030
        ThumbnailTransferFunction = &H5031
        ThumbnailSoftwareUsed = &H5032
        ThumbnailDateTime = &H5033
        ThumbnailArtist = &H5034
        ThumbnailWhitePoint = &H5035
        ThumbnailPrimaryChromaticities = &H5036
        ThumbnailYCbCrCoefficients = &H5037
        ThumbnailYCbCrSubsampling = &H5038
        ThumbnailYCbCrPositioning = &H5039
        ThumbnailRefBlackWhite = &H503A
        ThumbnailCopyRight = &H503B
        LuminanceTable = &H5090
        ChrominanceTable = &H5091
        FrameDelay = &H5100
        LoopCount = &H5101
        PixelUnit = &H5110
        PixelPerUnitX = &H5111
        PixelPerUnitY = &H5112
        PaletteHistogram = &H5113
        ExifExposureTime = &H829A
        ExifFNumber = &H829D
        ExifExposureProg = &H8822
        ExifSpectralSense = &H8824
        ExifISOSpeed = &H8827
        ExifOECF = &H8828
        ExifVer = &H9000
        ExifDTOrig = &H9003
        ExifDTDigitized = &H9004
        ExifCompConfig = &H9101
        ExifCompBPP = &H9102
        ExifShutterSpeed = &H9201
        ExifAperture = &H9202
        ExifBrightness = &H9203
        ExifExposureBias = &H9204
        ExifMaxAperture = &H9205
        ExifSubjectDist = &H9206
        ExifMeteringMode = &H9207
        ExifLightSource = &H9208
        ExifFlash = &H9209
        ExifFocalLength = &H920A
        ExifMakerNote = &H927C
        ExifUserComment = &H9286
        ExifDTSubsec = &H9290
        ExifDTOrigSS = &H9291
        ExifDTDigSS = &H9292
        ExifFPXVer = &HA000
        ExifColorSpace = &HA001
        ExifPixXDim = &HA002
        ExifPixYDim = &HA003
        ExifRelatedWav = &HA004
        ExifInterop = &HA005
        ExifFlashEnergy = &HA20B
        ExifSpatialFR = &HA20C
        ExifFocalXRes = &HA20E
        ExifFocalYRes = &HA20F
        ExifFocalResUnit = &HA210
        ExifSubjectLoc = &HA214
        ExifExposureIndex = &HA215
        ExifSensingMethod = &HA217
        ExifFileSource = &HA300
        ExifSceneType = &HA301
        ExifCfaPattern = &HA302
        GpsVer = &H0
        GpsLatitudeRef = &H1
        GpsLatitude = &H2
        GpsLongitudeRef = &H3
        GpsLongitude = &H4
        GpsAltitudeRef = &H5
        GpsAltitude = &H6
        GpsGpsTime = &H7
        GpsGpsSatellites = &H8
        GpsGpsStatus = &H9
        GpsGpsMeasureMode = &HA
        GpsGpsDop = &HB
        GpsSpeedRef = &HC
        GpsSpeed = &HD
        GpsTrackRef = &HE
        GpsTrack = &HF
        GpsImgDirRef = &H10
        GpsImgDir = &H11
        GpsMapDatum = &H12
        GpsDestLatRef = &H13
        GpsDestLat = &H14
        GpsDestLongRef = &H15
        GpsDestLong = &H16
        GpsDestBearRef = &H17
        GpsDestBear = &H18
        GpsDestDistRef = &H19
        GpsDestDist = &H1A
    End Enum


    Public Enum ExifDataTypes As Short
        UnsignedByte = 1
        AsciiString = 2
        UnsignedShort = 3
        UnsignedLong = 4
        UnsignedRational = 5
        SignedByte = 6
        Undefined = 7
        SignedShort = 8
        SignedLong = 9
        SignedRational = 10
        SingleFloat = 11
        DoubleFloat = 12
    End Enum

    '************************************************************************************************************
    '************************************************************************************************************
    '* Retrieve time when image was taken (EXIF DateTimeOriginal).
    '*

    Public Function GetPropertyString(ByVal PID As Int32, Optional ByVal DefaultValue As String = "") As String
        If Me.IsPropertyDefined(PID) Then
            Return GetString(Me._Image.GetPropertyItem(PID).Value)
        Else
            Return DefaultValue
        End If
    End Function

    Public Function IsPropertyDefined(ByVal PID As Int32) As Boolean
        Return CBool([Array].IndexOf(Me._Image.PropertyIdList, PID) > -1)
    End Function

    Private Function GetString(ByVal B As Byte()) As String
        Dim R As String = Me._Encoding.GetString(B)
        If R.EndsWith(vbNullChar) Then R = R.Substring(0, R.Length - 1)
        Return R
    End Function

    Public Sub SetProperty(ByVal PID As Int32, ByVal Data() As Byte, ByVal Type As ExifDataTypes)
        Dim P As System.Drawing.Imaging.PropertyItem = Me._Image.PropertyItems(0)
        P.Id = PID
        P.Value = Data
        P.Type = Type
        P.Len = Data.Length
        Me._Image.SetPropertyItem(P)
    End Sub


    Public Sub SetPropertyString(ByVal PID As Int32, ByVal Value As String)
        Dim Data() As Byte = Me._Encoding.GetBytes(Value & vbNullChar)
        SetProperty(PID, Data, ExifDataTypes.AsciiString)
    End Sub

    Public Property DateTimeOriginal() As DateTime
        Get
            Try
                Return DateTime.ParseExact(Me.GetPropertyString(TagNames.ExifDTOrig), "yyyy\:MM\:dd HH\:mm\:ss", Nothing)
            Catch ex As Exception
                Return DateTime.MinValue
            End Try
        End Get
        Set(ByVal Value As DateTime)
            Try
                Me.SetPropertyString(TagNames.ExifDTOrig, Value.ToString("yyyy\:MM\:dd HH\:mm\:ss"))
            Catch ex As Exception
            End Try
        End Set
    End Property





End Class
