#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: Acrobat_API_Core_UDF
    Description...........: UDF for Adobe Acrobat objects API core functions
	Dependencies..........: Adobe Acrobat API
	
	Documentation.........: https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/iac_api_reference.pdf
							https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/iac_developer_guide.pdf
							https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/acrobat_pdfl_api_reference.pdf
							https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/js_api_reference.pdf

	Credits...............: seadoggie01 @ AutoIt (https://www.autoitscript.com/forum/profile/107750-seadoggie01/)
    Author................: exorcistas@github.com
    Modified..............: 2020-08-02
    Version...............: v0.8.9.10b
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once

#Region GLOBAL_VARIABLES

	Global $__ACROBAT_DEBUG = True

	;-- PDSaveOptions
	Global Enum $__ACROBAT_SaveIncremental = 0, _
				$__ACROBAT_SaveFull = 1, _
				$__ACROBAT_SaveCopy = 2, _
				$__ACROBAT_SaveLinearized = 4, _
				$__ACROBAT_SaveWithPSHeader = 8, _
				$__ACROBAT_SaveBinaryOK = 10, _
				$__ACROBAT_SaveCollectGarbage = 20, _
				$__ACROBAT_SaveForceIncremental = 40, _
				$__ACROBAT_SaveKeepModDate = 80, _
				$__ACROBAT_SaveLeaveOpen = 100

	;-- PageViewMode
	Global Enum $__ACROBAT_DontCare = 0, _
				$__ACROBAT_UseNone, _
				$__ACROBAT_UseThumbs, _
				$__ACROBAT_UseBookmarks, _
				$__ACROBAT_FullScreen

#EndRegion GLOBAL_VARIABLES

#Region ACROBAT_OBJECTS
#cs	===================================================================================================================================
	AcroExch.App
	AcroExch.AVDoc
	AcroExch.PDDoc
	AcroExch.HiliteList (unmapped)
	AcroExch.AVPageView
	AcroExch.PDPage
	AcroExch.PDAnnot (unmapped)
	AcroExch.PDBookmark (unmapped)
	AcroExch.PDTextSelect
	AxAcroPDFLib.AxAcroPDF (unmapped)
	AcroExch.Point
	AcroExch.Rect
	AcroExch.Time (unmapped)
#ce	===================================================================================================================================
#EndRegion ACROBAT_OBJECTS

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
%% AcroExch.App	%%				The application itself
	_AcroExch_App_Create()
	_AcroExch_App_Show($_oAcroApp)
	_AcroExch_App_Hide($_oAcroApp)
	_AcroExch_App_Minimize($_oAcroApp)
	_AcroExch_App_Maximize($_oAcroApp)
	_AcroExch_App_Restore($_oAcroApp)
	_AcroExch_App_GetNumAVDocs($_oAcroApp)
	_AcroExch_App_GetActiveDoc($_oAcroApp)
	_AcroExch_App_GetAVDoc($_oAcroApp, $_iDocIndex)
	_AcroExch_App_CloseAllDocs($_oAcroApp)
	_AcroExch_App_Exit($_oAcroApp)

%% AcroExch.AVDoc %%			A document as seen in the user interface
	_AcroExch_AVDoc_Create()
	_AcroExch_AVDoc_Open($_oAcroAVDoc, $_sFullFilePath, $_sTempTitle = "")
	_AcroExch_AVDoc_Close($_oAcroAVDoc, $_bNoSave = True)
	_AcroExch_AVDoc_BringToFront($_oAcroAVDoc)
	_AcroExch_AVDoc_GetViewMode($_oAcroAVDoc)
	_AcroExch_AVDoc_SetViewMode($_oAcroAVDoc, $_iType = $__ACROBAT_DontCare)
	_AcroExch_AVDoc_FindText($_oAcroAVDoc, $_sText, $_bCaseSensitive = False, $_bReset = True)
	_AcroExch_AVDoc_GetAVPageView($_oAcroAVDoc)
	_AcroExch_AVDoc_GetPDDoc($_oAcroAVDoc)

%% AcroExch.AVPageView %%		The area of the window that displays the contents of a page
	_AcroExch_AVPageView_GetDoc($_oAcroAVPageView)
	_AcroExch_AVPageView_GetAVDoc($_oAcroAVPageView)
	_AcroExch_AVPageView_GetPage($_oAcroAVPageView)
	_AcroExch_AVPageView_GoTo($_oAcroAVPageView, $_iPage = 0)

%% AcroExch.PDDoc %%			The underlying PDF representation of a document
	_AcroExch_PDDoc_Create()
	_AcroExch_PDDoc_Open($_oAcroPDDoc, $_sFullFilePath)
	_AcroExch_PDDoc_OpenAVDoc($_oAcroPDDoc, $_sTitle = "")
	_AcroExch_PDDoc_Close($_oAcroPDDoc)
	_AcroExch_PDDoc_Save($_oAcroPDDoc, $_sSaveFullPath, $_iType = $__ACROBAT_SaveFull + $__ACROBAT_SaveCollectGarbage)
	_AcroExch_PDDoc_GetFileName($_oAcroPDDoc)
	_AcroExch_PDDoc_GetNumPages($_oAcroPDDoc)
	_AcroExch_PDDoc_AcquirePage($_oAcroPDDoc, $_iPageIndex = 0)
	_AcroExch_PDDoc_InsertPages($_oAcroPDDoc, $_oAcroPDDocSource, $_iSourceNumPages, $_iSourceStartPage = 0, $_iInsertPageAfter = 0, $_bCopyBookmarks = False)
	_AcroExch_PDDoc_MovePage($_oAcroPDDoc, $_iPageToMove, $_iMoveAfterThisPage = 0)
	_AcroExch_PDDoc_DeletePages($_oAcroPDDoc, $_iStartPage = 0, $_iEndPage = 0)
	_AcroExch_PDDoc_CreateTextSelect($_oAcroPDDoc, $_oAcroRect, $_iSelectPage = 0)
	_AcroExch_PDDoc_GetJSObject($_oAcroPDDoc)

%% AcroExch.PDPage %%			A single page in the PDF representation of a document
	_AcroExch_PDPage_GetDoc($_oAcroPDPage)
	_AcroExch_PDPage_GetNumber($_oAcroPDPage)
	_AcroExch_PDPage_GetSize($_oAcroPDPage)
	_AcroExch_PDPage_GetRotate($_oAcroPDPage)
	_AcroExch_PDPage_SetRotate($_oAcroPDPage, $_iRotateAngle = 90)

%% AcroExch.PDTextSelect %%		A selection of text on a single page
	_AcroExch_PDTextSelect_GetNumText($_oAcroPDTextSelect)
	_AcroExch_PDTextSelect_GetText($_oAcroPDTextSelect, $_iNumText)
	_AcroExch_PDTextSelect_Destroy($_oAcroPDTextSelect)

%% AcroExch.Rect %%				A rectange, specified by the top-left and bottom-right points
	_AcroExch_Rect_GetBottom($_oAcroRect)
	_AcroExch_Rect_GetLeft($_oAcroRect)
	_AcroExch_Rect_GetRight($_oAcroRect)
	_AcroExch_Rect_GetTop($_oAcroRect)
	_AcroExch_Rect_SetBottom($_oAcroRect, $_iValue)
	_AcroExch_Rect_SetLeft($_oAcroRect, $_iValue)
	_AcroExch_Rect_SetRight($_oAcroRect, $_iValue)
	_AcroExch_Rect_SetTop($_oAcroRect, $_iValue)

%% AcroExch.Point %%			A point, specified by its x and y coordinates
	_AcroExch_Point_GetX($_oAcroPoint)
	_AcroExch_Point_GetY($_oAcroPoint)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region AcroExch.App

    #cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Create()
        Description ...: Creates Acrobat application object     
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: -
		Returns........: On success, returns Adobe Acrobat application object
						 On failure, @error is set, FALSE is returned
		
        Remarks........: If object already exists and is found, new object is not created
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Create()
		Local $_oAcroApp = ObjGet("", "AcroExch.App")
			If NOT @error Then Return $_oAcroApp

		$_oAcroApp = ObjCreate("AcroExch.App")
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Create]: " & $_oAcroApp & @CRLF & @CRLF)
			If NOT IsObj($_oAcroApp) Then Return SetError(2, @error, False)
	
		Return $_oAcroApp
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Show($_oAcroApp)
        Description ...: Displays Acrobat application 
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, @error is set, FALSE is returned

        Remarks........: Shows the Acrobat application. When the viewer is shown, the user is in control, and the Acrobat application
						 does not automatically exit when the last automation object is destroyed.
						 However, it will exit if no documents are being displayed.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Show($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.Show()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Show]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Hide($_oAcroApp)
        Description ...: Hides Acrobat application
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, @error is set, FALSE is returned

        Remarks........: Hides the Acrobat application. When the viewer is hidden, the user has no control over it, and the Acrobat application
						 automatically exits when the last automation object is closed.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Hide($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.Hide()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Hide]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Minimize($_oAcroApp)
        Description ...: Minimizes the Acrobat application
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, @error is set, FALSE is returned

        Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Minimize($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.Minimize(1)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Minimize]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Maximize($_oAcroApp)
        Description ...: Maximizes the Acrobat application
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, @error is set, FALSE is returned

        Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Maximize($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.Maximize(1)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Maximize]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Restore($_oAcroApp)
        Description ...: Restores Acrobat application view
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, @error is set, FALSE is returned

		Remarks........: Restores the main window of the Acrobat application. 
						 Calling this with input set to positive number causes the main window to be 
						 restored to its original size and position and to become active.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Restore($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.Restore(1)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Restore]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_GetNumAVDocs($_oAcroApp)
        Description ...: Counts number of open Acrobat AVDoc objects
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: Number of open AcroExch.AVDoc objects

		Remarks........: The maximum number of documents the Acrobat application can open at a time is specified by 
						 <avpMaxOpenDocuments> preference, which can be obtained with <App.GetPreferenceEx> and set by <App.SetPreferenceEx>
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_GetNumAVDocs($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.GetNumAVDocs()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_GetNumAVDocs]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_GetActiveDoc($_oAcroApp)
        Description ...: Gets last active document object
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: {$_oAcroPDDoc} - document object
						 On failure, sets @error, returns FALSE

		Remarks........: Gets the frontmost document
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_GetActiveDoc($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_oAcroPDDoc = $_oAcroApp.GetActiveDoc()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_GetActiveDoc]: " & $_oAcroPDDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDDoc) Then Return SetError(2, 0, False)

		Return $_oAcroPDDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_GetAVDoc($_oAcroApp, $_iDocIndex)
        Description ...: Gets document object by index
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: {$_oAcroPDDoc} - document object
						 On failure, sets @error, returns FALSE

		Remarks........: Gets an AVDoc object from its index within the list of open AVDoc objects.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_GetAVDoc($_oAcroApp, $_iDocIndex)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_oAcroPDDoc = $_oAcroApp.GetAVDoc($_iDocIndex)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_GetAVDoc]: " & $_oAcroPDDoc & @CRLF & _
												 "DocIndex:	" & $_iDocIndex & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDDoc) Then Return SetError(2, 0, False)

		Return $_oAcroPDDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_CloseAllDocs($_oAcroApp)
        Description ...: Closes all open documents
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE

		Remarks........: You must explicitly close all documents, otherwise the process never exits.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_CloseAllDocs($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.CloseAllDocs()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_CloseAllDocs]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_App_Exit($_oAcroApp)
        Description ...: Exits Acrobat
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: {$_oAcroApp} - main Acrobat object
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE

		Remarks........: NOTE: Use <App.CloseAllDocs> to close all documents before calling this method.
						 Method does not work if application is visible (user is in control of it) - use <App.Hide> before calling this.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_App_Exit($_oAcroApp)
		If NOT IsObj($_oAcroApp) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroApp.Exit()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_App_Exit]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

#EndRegion AcroExch.App

#Region AcroExch.AVDoc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_Create()
        Description ...: Creates AVDoc object for individual document
		Documentation..: See <iac_api_reference.pdf>
		
        Parameters.....: -
		Returns........: {$_oAcroAVDoc} - object for individual document
						 On failure, sets @error, returns FALSE

		Remarks........: A view of a PDF document in a window. There is one AVDoc object per displayed document.
						 Unlike a PDDoc object, an AVDoc object has a window associated with it.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_Create()
		Local $_oAcroAVDoc = ObjCreate("AcroExch.AVDoc")

			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_Create]: " & $_oAcroAVDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, @extended, False)

		Return $_oAcroAVDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_Open($_oAcroAVDoc, $_sFullFilePath, $_sTempTitle = "")
        Description ...: Opens a file
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
						 {$_sFullFilePath} - path to PDF document
						 {$_sTempTitle} - (optional) title to be displayed for document
        
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE

		Remarks........: A new instance of AVDoc must be created for each displayed PDF file.
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_Open($_oAcroAVDoc, $_sFullFilePath, $_sTempTitle = "")
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroAVDoc.Open($_sFullFilePath, $_sTempTitle)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_Open]: " & Number($_iRetVal) & @CRLF & _
												 "FullFilePath:	" & $_sFullFilePath & @CRLF & _
												 "TempTitle:	" & $_sTempTitle & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_Close($_oAcroAVDoc, $_bNoSave = True)
        Description ...: Closes a document
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
						 {$_bNoSave} - if TRUE, document closed without saving, otherwise if modified, user is prompted for choice
        
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_Close($_oAcroAVDoc, $_bNoSave = True)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		Local $_iSave = 1
			If $_bNoSave Then $_iSave = 0

		$_iRetVal = $_oAcroAVDoc.Close($_bNoSave)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_Close]: " & Number($_iRetVal) & @CRLF & _
												 "NoSave:	" & $_bNoSave & @CRLF & @CRLF)
			If $_iRetVal = -1 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_BringToFront($_oAcroAVDoc)
        Description ...: Brings the window to the front
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_BringToFront($_oAcroAVDoc)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroAVDoc.BringToFront()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_BringToFront]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_GetTitle($_oAcroAVDoc)
        Description ...: Gets the window's title
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
		Returns........: Window title of associated document object

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_GetTitle($_oAcroAVDoc)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_sRetVal = $_oAcroAVDoc.GetTitle()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_GetTitle]: " & $_sRetVal & @CRLF & @CRLF)
	
		Return $_sRetVal
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_GetViewMode($_oAcroAVDoc)
        Description ...: Gets the current document view mode
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
		Returns........: Window view mode of associated document object

		Remarks........: 0 - leave the view mode as it is
						 1 - display without bookmarks or thumbnails
						 2 - display using thumbnails
						 3 - display using bookmarks
						 4 - display in full screen mode
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_GetViewMode($_oAcroAVDoc)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroAVDoc.GetViewMode()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_GetViewMode]: " & Number($_iRetVal) & @CRLF & @CRLF)
	
		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_SetViewMode($_oAcroAVDoc, $_iType = $__ACROBAT_DontCare)
        Description ...: Sets the mode in which the document will be viewed
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
						 {$_iType} - integer (0-4) representing mode type. See remarks
        
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE

		Remarks........: 0 - leave the view mode as it is
						 1 - display without bookmarks or thumbnails
						 2 - display using thumbnails
						 3 - display using bookmarks
						 4 - display in full screen mode
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_SetViewMode($_oAcroAVDoc, $_iType = $__ACROBAT_DontCare)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroAVDoc.SetViewMode($_iType)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_SetViewMode]: " & Number($_iRetVal) & @CRLF & _
												 "Type:	" & $_iType & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_FindText($_oAcroAVDoc, $_sText, $_bCaseSensitive = False, $_bReset = True)
        Description ...: Finds the specified text, scrolls so that it is visible, and highlights it
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
						 {$_sText} - text to search
						 {$_bCaseSensitive} - search in case sensitive mode
						 {$_bReset} - start search from 1st page, if false, search from current page
        
		Returns........: On success (text found), returns TRUE
						 On failure (or text not found), sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_FindText($_oAcroAVDoc, $_sText, $_bCaseSensitive = False, $_bReset = True)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		Local $_iCase = 0
			If $_bCaseSensitive Then $_iCase = 1

		Local $_iPageStart = 0
			If $_bReset Then $_iPageStart = 1

		$_iRetVal = $_oAcroAVDoc.FindText($_sText, $_iCase, $_iPageStart)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_FindText]: " & Number($_iRetVal) & @CRLF & _
												 "Text:	" & $_sText & @CRLF & _
												 "CaseSensitive:	" & $_bCaseSensitive & @CRLF & _
												 "Reset:	" & $_bReset & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return False
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_GetAVPageView($_oAcroAVDoc)
        Description ...: Gets AVPageView object associated with AVDoc
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
		Returns........: AVPageView object
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_GetAVPageView($_oAcroAVDoc)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_oAcroAVPageView = $_oAcroAVDoc.GetAVPageView()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_GetAVPageView]: " & $_oAcroAVPageView & @CRLF & @CRLF)
			If NOT IsObj($_oAcroAVPageView) Then Return SetError(2, 0, False)
	
		Return $_oAcroAVPageView
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVDoc_GetPDDoc($_oAcroAVDoc)
        Description ...: Gets PDDoc object associated with AVDoc
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVDoc} - object to be associated with opened document
		Returns........: PDDoc object
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVDoc_GetPDDoc($_oAcroAVDoc)
		If NOT IsObj($_oAcroAVDoc) Then Return SetError(1, 0, False)

		$_oAcroPDDoc = $_oAcroAVDoc.GetPDDoc()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVDoc_GetPDDoc]: " & $_oAcroPDDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDDoc) Then Return SetError(2, 0, False)
	
		Return $_oAcroPDDoc
	EndFunc

#EndRegion AcroExch.AVDoc

#Region AcroExch.AVPageView

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVPageView_GetDoc($_oAcroAVPageView)
        Description ...: Gets PDDoc object corresponding to the current page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVPageView} - object to be associated with page of document
		Returns........: PDDoc object
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVPageView_GetDoc($_oAcroAVPageView)
		If NOT IsObj($_oAcroAVPageView) Then Return SetError(1, 0, False)
		
		$_oAcroPDDoc = $_oAcroAVPageView.GetDoc()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVPageView_GetDoc]: " & $_oAcroPDDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDDoc) Then Return SetError(2, 0, False)

		Return $_oAcroPDDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVPageView_GetAVDoc($_oAcroAVPageView)
        Description ...: Gets AVDoc object corresponding to the current page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVPageView} - object to be associated with page of document
		Returns........: AVDoc object
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVPageView_GetAVDoc($_oAcroAVPageView)
		If NOT IsObj($_oAcroAVPageView) Then Return SetError(1, 0, False)
		
		$_oAcroAVDoc = $_oAcroAVPageView.GetAVDoc()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVPageView_GetAVDoc]: " & $_oAcroAVDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroAVDoc) Then Return SetError(2, 0, False)

		Return $_oAcroAVDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVPageView_GetPage($_oAcroAVPageView)
        Description ...: Gets PDPage object corresponding to the current page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVPageView} - object to be associated with page of document
		Returns........: PDPage object
						 On failure, sets @error, returns FALSE

		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVPageView_GetPage($_oAcroAVPageView)
		If NOT IsObj($_oAcroAVPageView) Then Return SetError(1, 0, False)
		
		$_oAcroPDPage = $_oAcroAVPageView.GetPage()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVPageView_GetPage]: " & $_oAcroPDPage & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDPage) Then Return SetError(2, 0, False)

		Return $_oAcroPDPage
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_AVPageView_GoTo($_oAcroAVPageView, $_iPage = 0)
        Description ...: Goes to specified page                 
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroAVPageView} - object to be associated with page of document
						 {$_iPage} - page number (always starts from 0)
        
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_AVPageView_GoTo($_oAcroAVPageView, $_iPage = 0)
		If NOT IsObj($_oAcroAVPageView) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroAVPageView.GoTo($_iPage)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_AVPageView_GoTo]: " & Number($_iRetVal) & @CRLF & _
												 "Page:	" & $_iPage & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)

		Return True
	EndFunc

#EndRegion AcroExch.AVPageView

#Region AcroExch.PDDoc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_Create()
        Description ...: Creates PDDoc object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: -
		Returns........: PDDoc object
						 On failure, sets @error, returns FALSE
		
		Remarks........: The underlying PDF representation of a document. PDDoc object is the hidden object behind every AVDoc object.
						 Through PDDoc objects, your application can perform most of the Document menu items from Acrobat (i.e. retrieve information)
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_Create()
		Local $_oAcroPDDoc = ObjCreate("AcroExch.PDDoc")
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_Create]: " & $_oAcroPDDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDDoc) Then Return SetError(2, @extended, False)

		Return $_oAcroPDDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_Open($_oAcroPDDoc, $_sFullFilePath)
        Description ...: Opens a file
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object to associate file
						 {$_sFullFilePath} - full path to file to open

		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: New instance of PDDoc must be created for each open PDF file
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_Open($_oAcroPDDoc, $_sFullFilePath)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroPDDoc.Open($_sFullFilePath)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_Open]: " & Number($_iRetVal) & @CRLF & _
												 "FullFilePath:	" & $_sFullFilePath & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_OpenAVDoc($_oAcroPDDoc, $_sTitle = "")
        Description ...: Opens a window and displays document in it
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
						 {$_sTitle} - (optional) title display for document window

		Returns........: AVDoc object
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_OpenAVDoc($_oAcroPDDoc, $_sTitle = "")
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_oAcroAVDoc = $_oAcroPDDoc.OpenAVDoc($_sTitle)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_OpenAVDoc]: " & $_oAcroAVDoc & @CRLF & _
												 "Title:	" & $_sTitle & @CRLF & @CRLF)
			If NOT IsObj($_oAcroAVDoc) Then Return SetError(2, 0, False)

		Return $_oAcroAVDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_Close($_oAcroPDDoc)
        Description ...: Closes a file
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: If PDDoc and AVDoc are constructed with same file, this method destroys both objects
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_Close($_oAcroPDDoc)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroPDDoc.Close()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_Close]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_Save($_oAcroPDDoc, $_sSaveFullPath, $_iType = $__ACROBAT_SaveFull + $__ACROBAT_SaveCollectGarbage)
        Description ...: Saves a document
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
						 {$_sSaveFullPath} - full path where to save document
						 {$_iType} - (optional) integer to represent way to save document

		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: See documentation file for details
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_Save($_oAcroPDDoc, $_sSaveFullPath, $_iType = $__ACROBAT_SaveFull + $__ACROBAT_SaveCollectGarbage)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroPDDoc.Save($_iType, $_sSaveFullPath)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_Save]: " & Number($_iRetVal) & @CRLF & _
												 "SaveFullPath:	" & $_sSaveFullPath & @CRLF & _
												 "Type:	" & $_iType & @CRLF & @CRLF)
			If $_iRetVal = 0 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_GetFileName($_oAcroPDDoc)
        Description ...: Gets filename associated with PDDoc object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file

		Returns........: Filename
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_GetFileName($_oAcroPDDoc)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_sRetVal = $_oAcroPDDoc.GetFileName()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_GetFileName]: " & $_sRetVal & @CRLF & @CRLF)

		Return $_sRetVal
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_GetNumPages($_oAcroPDDoc)
        Description ...: Gets number of pages in a file
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file

		Returns........: Number of pages
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_GetNumPages($_oAcroPDDoc)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroPDDoc.GetNumPages()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_GetNumPages]: " & Number($_iRetVal) & @CRLF & @CRLF)
			If $_iRetVal = -1 Then Return SetError(2, 0, False)
	
		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_AcquirePage($_oAcroPDDoc, $_iPageIndex = 0)
        Description ...: Acquires specified page PDPage object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
						 {$_iPageIndex} - (optional) page index (always starts with 0)

		Returns........: PDPage object
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_AcquirePage($_oAcroPDDoc, $_iPageIndex = 0)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_oAcroPDPage = $_oAcroPDDoc.AcquirePage($_iPageIndex)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_AcquirePage]: " & $_oAcroPDPage & @CRLF & _
												 "PageIndex:	" & $_iPageIndex & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDPage) Then Return SetError(2, 0, False)

		Return $_oAcroPDPage
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_InsertPages($_oAcroPDDoc, $_oAcroPDDocSource, $_iSourceNumPages, $_iSourceStartPage = 0, $_iInsertPageAfter = 0, $_bCopyBookmarks = False)
        Description ...: Inserts specified pages from source document after indicated page with current document
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file (to insert into)
						 {$_oAcroPDDocSource} - PDDoc object of source file to be inserted
						 {$_iSourceNumPages} - number of pages to be inserted
						 {$_iSourceStartPage} - (optional) first page of source to be inserted
						 {$_iInsertPageAfter} - (optional) current document page to insert after
						 {$_bCopyBookmarks} - (optional) copy bookmarks from source


		Returns........: On success, return TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_InsertPages($_oAcroPDDoc, $_oAcroPDDocSource, $_iSourceNumPages, $_iSourceStartPage = 0, $_iInsertPageAfter = 0, $_bCopyBookmarks = False)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		Local $_iBookmarks = 0
		If $_bCopyBookmarks Then $_iBookmarks = 1

		$_iRetVal = $_oAcroPDDoc.InsertPages($_iInsertPageAfter, $_oAcroPDDocSource, $_iSourceStartPage, $_iSourceNumPages, $_iBookmarks)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_InsertPages]: " & Number($_iRetVal) & @CRLF & _
												 "PDDocSource:	" & $_oAcroPDDocSource & @CRLF & _
												 "SourceStartPage:	" & $_iSourceStartPage & @CRLF & _
												 "SourceNumPages:	" & $_iSourceNumPages & @CRLF & _
												 "InsertPageAfter:	" & $_iInsertPageAfter & @CRLF & _
												 "CopyBookmarks:	" & $_bCopyBookmarks & @CRLF & @CRLF)
			If $_iRetVal = -1 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_MovePage($_oAcroPDDoc, $_iPageToMove, $_iMoveAfterThisPage = 0)
        Description ...: Moves page to another location within same document
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
						 {$_iPageToMove} - index of page to be moved
						 {$_iMoveAfterThisPage} - number of page to be inserted after

		Returns........: On success, return TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_MovePage($_oAcroPDDoc, $_iPageToMove, $_iMoveAfterThisPage = 0)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)

		$_iRetVal = $_oAcroPDDoc.MovePage($_iMoveAfterThisPage, $_iPageToMove)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_MovePage]: " & Number($_iRetVal) & @CRLF & _
												 "MoveAfterThisPage:	" & $_iMoveAfterThisPage & @CRLF & _
												 "PageToMove:	" & $_iPageToMove & @CRLF & @CRLF)
			If $_iRetVal = -1 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_DeletePages($_oAcroPDDoc, $_iStartPage = 0, $_iEndPage = 0)
        Description ...: Deletes page from file
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
						 {$_iStartPage} - first page to be deleted
						 {$_iEndPage} - last page to be deleted

		Returns........: On success, return TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_DeletePages($_oAcroPDDoc, $_iStartPage = 0, $_iEndPage = 0)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPDDoc.InsertPages($_iStartPage, $_iEndPage)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_DeletePages]: " & Number($_iRetVal) & @CRLF & _
												 "StartPage:	" & $_iStartPage & @CRLF & _
												 "EndPage:	" & $_iEndPage & @CRLF & @CRLF)
			If $_iRetVal = -1 Then Return SetError(2, 0, False)
	
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_CreateTextSelect($_oAcroPDDoc, $_oAcroRect, $_iSelectPage = 0)
        Description ...: Creates text selection from specified rectangle on specified page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file
						 {$_oAcroRect} - first page to be deleted
						 {$_iSelectPage} - page index to use selection at

		Returns........: PDTextSelect object - selection object
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_CreateTextSelect($_oAcroPDDoc, $_oAcroRect, $_iSelectPage = 0)
		If NOT IsObj($_oAcroPDDoc) OR NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_oAcroPDTextSelect = $_oAcroPDDoc.CreateTextSelect($_iSelectPage, $_oAcroRect)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_CreateTextSelect]: " & $_oAcroPDTextSelect & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDTextSelect) Then Return SetError(2, 0, False)

		Return $_oAcroPDTextSelect
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDDoc_GetJSObject($_oAcroPDDoc)
        Description ...: Gets a dual interface to the JavaScript object associated with PDDoc
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDDoc} - PDDoc object associated with file

		Returns........: JS object
						 On failure, sets @error, returns FALSE
		
		Remarks........: This allows automation clients full access to both built-in and user-defined JS method available in document
						 See <js_api_reference.pdf>
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDDoc_GetJSObject($_oAcroPDDoc)
		If NOT IsObj($_oAcroPDDoc) Then Return SetError(1, 0, False)
		
		$_oAcroJS = $_oAcroPDDoc.GetJSObject()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDDoc_GetJSObject]: " & $_oAcroJS & @CRLF & @CRLF)
			If NOT IsObj($_oAcroJS) Then Return SetError(2, 0, False)

		Return $_oAcroJS
	EndFunc

#EndRegion AcroExch.PDDoc

#Region AcroExch.PDPage

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDPage_GetDoc($_oAcroPDPage)
        Description ...: Gets PDDoc object associated with page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDPage} - PDPage object associated with page

		Returns........: PDDoc object
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDPage_GetDoc($_oAcroPDPage)
		If NOT IsObj($_oAcroPDPage) Then Return SetError(1, 0, False)
		
		$_oAcroPDDoc = $_oAcroPDPage.GetDoc()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDPage_GetDoc]: " & $_oAcroPDDoc & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPDDoc) Then Return SetError(2, 0, False)

		Return $_oAcroPDDoc
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDPage_GetNumber($_oAcroPDPage)
        Description ...: Gets page number of PDPage object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDPage} - PDPage object associated with page

		Returns........: Number of current page
						 On failure, sets @error, returns FALSE
		
		Remarks........: First document page is always 0
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDPage_GetNumber($_oAcroPDPage)
		If NOT IsObj($_oAcroPDPage) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPDPage.GetNumber()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDPage_GetNumber]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDPage_GetSize($_oAcroPDPage)
        Description ...: Gets page's width and height in points (Point object)
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDPage} - PDPage object associated with page

		Returns........: Point object
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDPage_GetSize($_oAcroPDPage)
		If NOT IsObj($_oAcroPDPage) Then Return SetError(1, 0, False)
		
		$_oAcroPoint = $_oAcroPDPage.GetSize()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDPage_GetSize]: " & $_oAcroPoint & @CRLF & @CRLF)
			If NOT IsObj($_oAcroPoint) Then Return SetError(2, 0, False)

		Return $_oAcroPoint
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDPage_GetRotate($_oAcroPDPage)
        Description ...: Gets the rotation value in degrees for current page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDPage} - PDPage object associated with page

		Returns........: Rotation value in degrees
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDPage_GetRotate($_oAcroPDPage)
		If NOT IsObj($_oAcroPDPage) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPDPage.GetRotate()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDPage_GetRotate]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDPage_SetRotate($_oAcroPDPage, $_iRotateAngle = 90)
        Description ...: Sets the rotation value in degrees for current page
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDPage} - PDPage object associated with page
						 {$_iRotateAngle} - (optional) 0,90,180 or 270

		Returns........: On success, returns TRUE
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDPage_SetRotate($_oAcroPDPage, $_iRotateAngle = 90)
		If NOT IsObj($_oAcroPDPage) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPDPage.SetRotate($_iRotateAngle)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDPage_SetRotate]: " & Number($_iRetVal) & @CRLF & _
												  "RotateAngle:	" & $_iRotateAngle & @CRLF & @CRLF)
			If $_iRetVal = -1 Then Return SetError(2, 0, False)

		Return True
	EndFunc

#EndRegion AcroExch.PDPage

#Region AcroExch.PDTextSelect

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDTextSelect_GetNumText($_oAcroPDTextSelect)
        Description ...: Gets number of text elements in text selection
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDTextSelect} - PDTextSelect object associated with selection

		Returns........: Number of text elements in selection
						 On failure, sets @error, returns FALSE
		
		Remarks........: Use this method to determine how many times to call <GetText> to obtain all selection's text
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDTextSelect_GetNumText($_oAcroPDTextSelect)
		If NOT IsObj($_oAcroPDTextSelect) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPDTextSelect.GetNumText()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDTextSelect_GetNumText]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDTextSelect_GetText($_oAcroPDTextSelect, $_iNumText)
        Description ...: Gets text from specified element of a text selection
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDTextSelect} - PDTextSelect object associated with selection

		Returns........: Text
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDTextSelect_GetText($_oAcroPDTextSelect, $_iNumText)
		If NOT IsObj($_oAcroPDTextSelect) Then Return SetError(1, 0, False)
		
		$_sRetVal = $_oAcroPDTextSelect.GetText($_iNumText)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDTextSelect_GetText]: " & $_sRetVal & @CRLF & _
												 "NumText:	" & $_iNumText & @CRLF & @CRLF)

		Return $_sRetVal
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_PDTextSelect_Destroy($_oAcroPDTextSelect)
        Description ...: Destroys text selection object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPDTextSelect} - PDTextSelect object associated with selection

		Returns........: -
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_PDTextSelect_Destroy($_oAcroPDTextSelect)
		If NOT IsObj($_oAcroPDTextSelect) Then Return SetError(1, 0, False)
		
		$_oAcroPDTextSelect.Destroy()
		If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_PDTextSelect_Destroy]" & @CRLF & @CRLF)
	EndFunc

#EndRegion AcroExch.PDTextSelect

#Region AcroExch.Rect

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Rect_Create()
        Description ...: Creates Rect object to hold coordinates
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: -
		Returns........: Rect object
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Rect_Create()
		Local $_oAcroRect = ObjCreate("AcroExch.Rect")
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_Create]: " & $_oAcroRect & @CRLF & @CRLF)
			If NOT IsObj($_oAcroRect) Then Return SetError(1, @extended, False)

		Return $_oAcroRect
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Rect_GetBottom($_oAcroRect)
        Description ...: Gets y-coordinate of AcroRect object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroRect} - Rect object
		Returns........: Coordinate
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Rect_GetBottom($_oAcroRect)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroRect.Bottom()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_GetBottom]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

		#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Rect_GetLeft($_oAcroRect)
        Description ...: Gets x-coordinate of AcroRect object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroRect} - Rect object
		Returns........: Coordinate
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Rect_GetLeft($_oAcroRect)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroRect.Left()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_GetLeft]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Rect_GetRight($_oAcroRect)
        Description ...: Gets x-coordinate of AcroRect object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroRect} - Rect object
		Returns........: Coordinate
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Rect_GetRight($_oAcroRect)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroRect.Right()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_GetRight]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Rect_GetTop($_oAcroRect)
        Description ...: Gets x-coordinate of AcroRect object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroRect} - Rect object
		Returns........: Coordinate
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Rect_GetTop($_oAcroRect)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroRect.Top()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_GetTop]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	Func _AcroExch_Rect_SetBottom($_oAcroRect, $_iValue)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_oAcroRect.Bottom() = Number($_iValue)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_SetBottom]: " & Number($_iValue) & @CRLF & @CRLF)
	EndFunc

	Func _AcroExch_Rect_SetLeft($_oAcroRect, $_iValue)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_oAcroRect.Left() = Number($_iValue)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_SetLeft]: " & Number($_iValue) & @CRLF & @CRLF)
	EndFunc

	Func _AcroExch_Rect_SetRight($_oAcroRect, $_iValue)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_oAcroRect.Right() = Number($_iValue)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_SetRight]: " & Number($_iValue) & @CRLF & @CRLF)
	EndFunc

	Func _AcroExch_Rect_SetTop($_oAcroRect, $_iValue)
		If NOT IsObj($_oAcroRect) Then Return SetError(1, 0, False)
		
		$_oAcroRect.Top() = Number($_iValue)
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Rect_SetTop]: " & Number($_iValue) & @CRLF & @CRLF)
	EndFunc

#EndRegion AcroExch.Rect

#Region AcroExch.Point

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Point_GetX($_oAcroPoint)
        Description ...: Gets x-coordinate of AcroRect object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPoint} - Point object
		Returns........: Coordinate
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Point_GetX($_oAcroPoint)
		If NOT IsObj($_oAcroPoint) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPoint.X()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Point_GetX]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
        Name...........: _AcroExch_Point_GetY($_oAcroPoint)
        Description ...: Gets x-coordinate of AcroRect object
		Documentation..: See <iac_api_reference.pdf>
		
		Parameters.....: {$_oAcroPoint} - Point object
		Returns........: Coordinate
						 On failure, sets @error, returns FALSE
		
		Remarks........: -
                         
        Author ........: exorcistas@github.com
        Modified.......: 2020-07-31
    #ce ===============================================================================================================================
	Func _AcroExch_Point_GetY($_oAcroPoint)
		If NOT IsObj($_oAcroPoint) Then Return SetError(1, 0, False)
		
		$_iRetVal = $_oAcroPoint.Y()
			If $__ACROBAT_DEBUG Then ConsoleWrite("[_AcroExch_Point_GetY]: " & Number($_iRetVal) & @CRLF & @CRLF)

		Return Number($_iRetVal)
	EndFunc

#EndRegion AcroExch.Point