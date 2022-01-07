#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Name..................: Acrobat_API_Extended_UDF
    Description...........: UDF for Adobe Acrobat objects API extended functions
	Dependencies..........: Adobe Acrobat; <Acrobat_API_UDF_Core.au3>
	
	Documentation.........: https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/iac_api_reference.pdf
							https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/iac_developer_guide.pdf
							https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/acrobat_pdfl_api_reference.pdf
							https://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/js_api_reference.pdf

    Author................: exorcistas@github.com
    Modified..............: 2020-09-24
    Version...............: v0.1.92a
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once
#include <Array.au3>
#include <Acrobat_API_UDF_Core.au3>

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
%% APPLICATION %%
	_Acro_AppShutdown()

%% DOCUMENT %%
	_Acro_DocOpen($_sFullFilePath)
	_Acro_DocGetText($_oAcroAVDoc, $_bReturnArray = False)
	_Acro_DocClose($_oAcroAVDoc, $_bNoSave = True)

%% PAGE %%
	_Acro_PageGetText($_oAcroAVDoc, $_iPageNumber)
	_Acro_PageGetCount($_oAcroAVDoc)

%% INTERNAL %%
	__Acro_PageTextSelect($_oAcroPDDoc, $_iPageNumber)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region APPLICATION_FUNCTIONS

	Func _Acro_AppShutdown()
		Local $_oAcroApp = _AcroExch_App_Create()
			If @error Then Return SetError(1, @error, False)
		_AcroExch_App_CloseAllDocs($_oAcroApp)
			If @error Then Return SetError(2, @error, False)
		_AcroExch_App_Hide($_oAcroApp)
			If @error Then Return SetError(3, @error, False)
		_AcroExch_App_Exit($_oAcroApp)
			If @error Then Return SetError(4, @error, False)
		
		Return True
	EndFunc

#EndRegion APPLICATION_FUNCTIONS

#Region DOCUMENT_FUNCTIONS

	Func _Acro_DocOpen($_sFullFilePath, $_bVisible = False)
		Local $_oAcroAVDoc = _AcroExch_AVDoc_Create()
			If @error Then Return SetError(1, @error, False)
		_AcroExch_AVDoc_Open($_oAcroAVDoc, $_sFullFilePath)
			If @error Then Return SetError(2, @error, False)
			
		If $_bVisible Then 
			_AcroExch_AVDoc_BringToFront($_oAcroAVDoc)
			If @error Then Return SetError(3, @error, False)
		EndIf

		Return $_oAcroAVDoc
	EndFunc

	Func _Acro_DocGetText($_oAcroAVDoc, $_bReturnArray = False)
		Local $_iPageCount = _Acro_PageGetCount($_oAcroAVDoc)
			If @error Then Return SetError(1, @error, False)
		
		Local $_aText[0]
		Local $_sText = ""
		Local $_sTemp = ""
		For $i = 0 To $_iPageCount-1
			$_sTemp = _Acro_PageGetText($_oAcroAVDoc, $i)
				If @error Then Return SetError(2, @error, False)

			If NOT $_bReturnArray Then
				$_sText &= $_sTemp
			Else
				_ArrayAdd($_aText, $_sTemp)
					If @error Then Return SetError(3, @error, False)
			EndIf
		Next
		
		Local $_vText = ($_bReturnArray) ? $_aText : $_sText

		Return $_vText
	EndFunc

	Func _Acro_DocTextExists($_oAcroAVDoc, $_sTextToFind, $_bCaseSensitive = False)
		Local $_bTextExists = _AcroExch_AVDoc_FindText($_oAcroAVDoc, $_sTextToFind, $_bCaseSensitive, True)
			If @error Then Return SetError(1, @error, False)

		Return $_bTextExists
	EndFunc

	Func _Acro_DocClose($_oAcroAVDoc, $_bNoSave = True)
		_AcroExch_AVDoc_Close($_oAcroAVDoc, $_bNoSave)
			If @error Then Return SetError(1, @error, False)

		Return True
	EndFunc

#EndRegion DOCUMENT_FUNCTIONS

#Region PAGE_FUNCTIONS

	Func _Acro_PageGoTo($_oAcroAVDoc, $_iPageNumber)
		Local $_oAcroAVPageView = _AcroExch_AVDoc_GetAVPageView($_oAcroAVDoc)
			If @error Then Return SetError(1, @error, False)

		_AcroExch_AVPageView_GoTo($_oAcroAVPageView, $_iPageNumber)
			If @error Then Return SetError(2, @error, False)
	EndFunc

	Func _Acro_PageGetText($_oAcroAVDoc, $_iPageNumber)
		Local $_oAcroPDDoc = _AcroExch_AVDoc_GetPDDoc($_oAcroAVDoc)
			If @error Then Return SetError(1, @error, False)
		Local $_oAcroPDTextSelect = __Acro_PageTextSelect($_oAcroPDDoc, $_iPageNumber)
			If @error Then Return SetError(2, @error, False)
		Local $_iNumText = _AcroExch_PDTextSelect_GetNumText($_oAcroPDTextSelect)
			If @error Then Return SetError(3, @error, False)

		Local $_sPageText = ""
		For $i = 0 To $_iNumText-1
			$_sPageText &= _AcroExch_PDTextSelect_GetText($_oAcroPDTextSelect, $i)
				If @error Then Return SetError(4, @error, False)
		Next
		_AcroExch_PDTextSelect_Destroy($_oAcroPDTextSelect)

		Return $_sPageText
	EndFunc

	Func _Acro_PageGetCount($_oAcroAVDoc)
		Local $_oAcroPDDoc = _AcroExch_AVDoc_GetPDDoc($_oAcroAVDoc)
			If @error Then Return SetError(1, @error, False)
		Local $_iPageCount = _AcroExch_PDDoc_GetNumPages($_oAcroPDDoc)
			If @error Then Return SetError(2, @error, False)

		Return $_iPageCount
	EndFunc

#EndRegion PAGE_FUNCTIONS

#Region INTERNAL_FUNCTIONS

	Func __Acro_PageTextSelect($_oAcroPDDoc, $_iPageNumber)
		$_oAcroPDPage = _AcroExch_PDDoc_AcquirePage($_oAcroPDDoc, $_iPageNumber)
			If @error Then Return SetError(1, @error, False)

		Local $_oAcroPoint = _AcroExch_PDPage_GetSize($_oAcroPDPage)
			If @error Then Return SetError(2, @error, False)

		Local $_oAcroRect = _AcroExch_Rect_Create()
			If @error Then Return SetError(3, @error, False)
		_AcroExch_Rect_SetBottom($_oAcroRect, 0)
		_AcroExch_Rect_SetTop($_oAcroRect, _AcroExch_Point_GetY($_oAcroPoint))
		_AcroExch_Rect_SetLeft($_oAcroRect, 0)
		_AcroExch_Rect_SetRight($_oAcroRect, _AcroExch_Point_GetX($_oAcroPoint))

		Local $_oAcroPDTextSelect =_AcroExch_PDDoc_CreateTextSelect($_oAcroPDDoc, $_oAcroRect, $_iPageNumber)
			If @error Then Return SetError(4, @error, False)
		
		Return $_oAcroPDTextSelect
	EndFunc

#EndRegion INTERNAL_FUNCTIONS