/**
 * XMLParser.cpp
 * 
 * Copyright (c) 2008 by <your name/ organization here>
 */
// XMLParser.cpp : DLL �A�v���P�[�V�����p�ɃG�N�X�|�[�g�����֐����`���܂��B
//

// http://msdn.microsoft.com/en-us/library/ms753794(v=vs.85).aspx
// http://hp.vector.co.jp/authors/VA014436/prg_memo/windows/vctips/006.html

#include "stdafx.h"
#include "XMLParser.h"
#include "SharedConstants.h"

// Macro that calls a COM method returning HRESULT value.
#define CHECK_HR(stmt)	do { hr=(stmt); if(hr!=S_OK) goto Cleanup; } while (0)

// Macro to verify memory allocation.
#define CHECK_ALLOC(p)	do { if(!(p)) { hr=E_OUTOFMEMORY; goto Cleanup; } } while (0)

// Macro that releases a COM object if not NULL.
#define SAFE_RELEASE(p)	do { if((p)) { (p)->Release(); (p)=NULL; } } while (0)


//static HRESULT VariantFromString(PCWSTR wszValue, VARIANT& Variant);

static BOOL LoadDOMRaw(IXMLDOMElement** pRootElement, PCTSTR szXMLuri);
static HRESULT CreateAndInitDOM(IXMLDOMDocument** ppDoc);
static void ReportError(PCTSTR szMessage);
static BOOL StoreValues(IXMLDOMNode* pNode, PCTSTR attrName, map<PCTSTR, PCTSTR>* keysMap);
static void Store(IXMLDOMNode* pNode, PCTSTR szAttrName, map<PCTSTR, PCTSTR>* keysMap);

//static LONG dump_exception(_EXCEPTION_POINTERS *ep);

static int g_COMInitializeCount = 0;

TCHAR szTitle[] = _T("XMLParser");

/**
 * @brief
 * COM�̏�������XML�̃��[�h���s���܂��B
 * 
 * @param pRootElement
 * IXMLDOMElement���i�[����|�C���^�̃|�C���^�B
 * 
 * @param szXMLuri
 * ���[�h����XML�t�@�C���̃p�X�B
 * 
 * @returns
 * ����������у��[�h�����������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * COM�̏�������XML�̃��[�h���s���܂��B
 * 
 * @remarks
 * �֐������������ꍇ�ɂ́A�v���O�����I���O��UnInitialize()���Ăяo���Ă��������B
 * 
 * @see
 * LoadDOMRaw() | UnInitialize()
 */
BOOL WINAPI Initialize(IXMLDOMElement** pRootElement, PCTSTR szXMLuri)
{
	BOOL bRet = TRUE;

	if (!g_COMInitializeCount) {
		if (FAILED(CoInitialize(NULL))) {
			return FALSE;
		}
		g_COMInitializeCount++;
	}
	
	if ((bRet = LoadDOMRaw(pRootElement, szXMLuri)) == FALSE) {
		if (g_COMInitializeCount > 0) {
			CoUninitialize();
			g_COMInitializeCount--;
		}
	}
	return bRet;
}

/**
 * @brief
 * COM�̏I��������IXMLDOMElement�̉�����s���܂��B
 * 
 * @param pRootElement
 * Initialize�ɂ���Ċi�[���ꂽIXMLDOMElement�|�C���^�B
 * 
 * COM�̏I��������IXMLDOMElement�̉�����s���܂��B
 * 
 * @remarks
 * Initialize�ɂ���Ċ֐������������ꍇ�ɂ́A�v���O�����I���O�ɖ{�֐����Ăяo���Ă��������B
 * 
 * @see
 * Initialize()
 */
void WINAPI UnInitialize(IXMLDOMElement* pRootElement)
{
	SAFE_RELEASE(pRootElement);
	if (g_COMInitializeCount > 0) {
		CoUninitialize();
		g_COMInitializeCount--;
	}
}

// �����������[�N����\��������̂Ŏg�p�֎~�B
// SysFreeString���ĂԕK�v������B
// Helper function to create a VT_BSTR variant from a null terminated string.
//HRESULT VariantFromString(PCWSTR wszValue, VARIANT& Variant)
//{
//	HRESULT hr = S_OK;
//	BSTR bstr = SysAllocString(wszValue);
//	CHECK_ALLOC(bstr);
//
//	V_VT(&Variant) = VT_BSTR;
//	V_BSTR(&Variant) = bstr;
//
//Cleanup:
//	return hr;
//}

/**
 * @brief
 * DOM�I�u�W�F�N�g�̐����Ƃ̏��������s���܂��B
 * 
 * @param ppDoc
 * IXMLDOMDocument�̃|�C���^�̃|�C���^�B
 * 
 * @returns
 * �������ʂ�Ԃ��܂��B
 * 
 * DOM�I�u�W�F�N�g�̐����Ƃ̏��������s���܂��B
 * 
 */
HRESULT CreateAndInitDOM(IXMLDOMDocument** ppDoc)
{
	HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(ppDoc));
	if (SUCCEEDED(hr)) {
		// these methods should not fail so don't inspect result.
		(*ppDoc)->put_async(VARIANT_FALSE);
		(*ppDoc)->put_validateOnParse(VARIANT_FALSE);
		(*ppDoc)->put_resolveExternals(VARIANT_FALSE);
	}
	return hr;
}

/**
 * @brief
 * XML�����[�h���AIXMLDOMElement�I�u�W�F�N�g�Ɋi�[���܂��B
 * 
 * @param pRootElement
 * XML���̃��[�g�G�������g���i�[����I�u�W�F�N�g�̃|�C���^�̃|�C���^�B
 * 
 * @param szXMLuri
 * ���[�h����XML�t�@�C���̃p�X�B
 * 
 * @returns
 * ���[�h����уI�u�W�F�N�g���ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * XML�����[�h���AIXMLDOMElement�I�u�W�F�N�g�Ɋi�[���܂��B
 */
BOOL LoadDOMRaw(IXMLDOMElement** pRootElement, PCTSTR szXMLuri)
{
	HRESULT hr = S_OK;

	IXMLDOMDocument* pXMLDom = NULL;
	IXMLDOMParseError* pXMLError = NULL;

	BSTR bstrXMLuri = NULL;
	BSTR bstrXML = NULL;
	BSTR bstrError = NULL;
	VARIANT_BOOL varStatus = 0;
	VARIANT varFileName = {0};
	VariantInit(&varFileName);

	CHECK_HR(CreateAndInitDOM(&pXMLDom));

	// XML file name to load
	bstrXMLuri = SysAllocString(szXMLuri);
	CHECK_ALLOC(bstrXMLuri);
	V_VT(&varFileName) = VT_BSTR;
	V_BSTR(&varFileName) = bstrXMLuri;

	CHECK_HR(pXMLDom->load(varFileName, &varStatus));
	if (varStatus == VARIANT_TRUE) {
		CHECK_HR(pXMLDom->get_xml(&bstrXML));

		// ���[�g�G�������g�擾
		*pRootElement = NULL;
		CHECK_HR(pXMLDom->get_documentElement(pRootElement));
		
	} else {
		CHECK_HR(pXMLDom->get_parseError(&pXMLError));
		CHECK_HR(pXMLError->get_reason(&bstrError));
		//ReportError(bstrError);
		hr = S_FALSE;
	}

Cleanup:
	SAFE_RELEASE(pXMLDom);
	SAFE_RELEASE(pXMLError);
	SysFreeString(bstrXMLuri);
	SysFreeString(bstrXML);
	SysFreeString(bstrError);
	VariantClear(&varFileName);

	return hr==S_OK;
}

/**
 * @brief
 * �w�肵�����[�g�G�������g����m�[�h�I�u�W�F�N�g���擾���܂��B
 * 
 * @param pRootElement
 * ���[�h����XML���i�[���ꂽ���[�g�G�������g�B
 * 
 * @param szNodeName
 * �m�[�h�I�u�W�F�N�g�Ώۂ̃m�[�h[�^�O]���B
 * 
 * @param pNode
 * �i�[����IXMLDOMNode�̃|�C���^�̃|�C���^�B
 * 
 * @returns
 * �m�[�h�I�u�W�F�N�g�̎擾�ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * �w�肵�����[�g�G�������g����A�w��m�[�h�̃m�[�h�I�u�W�F�N�g���擾���܂��B
 * 
 * @remarks
 * �g�p���FreeDOMTree()���Ăяo���Ă��������B
 * 
 * @see
 * FreeDOMTree()
 */
BOOL WINAPI GetDOMTree(IXMLDOMElement* pRootElement, PCTSTR szNodeName, IXMLDOMNode** pNode)
{
	HRESULT hr = S_OK;
	if (pRootElement == NULL) {
		//ReportError(TEXT("���XML�̓ǂݍ��݂��s���Ă��������B"));
		return FALSE;
	}
	BSTR bstrNodeName = NULL;
	bstrNodeName = SysAllocString(szNodeName);
	CHECK_ALLOC(bstrNodeName);
	CHECK_HR(pRootElement->selectSingleNode(bstrNodeName, pNode));

Cleanup:
	SysFreeString(bstrNodeName);
	return hr==S_OK;
}

/**
 * @brief
 * �m�[�h�I�u�W�F�N�g�̉�����s���܂��B
 * 
 * @param pNode
 * IXMLDOMNode�̃|�C���^�B
 * 
 * GetDOMTree()�Ŏ擾�����m�[�h�I�u�W�F�N�g�̉�����s���܂��B
 * 
 * @remarks
 * GetDOMTree()�Ŏ擾�����m�[�h�I�u�W�F�N�g�͖{�֐��ŉ�����s���܂��B
 * 
 * @see
 * GetDOMTree()
 */
void WINAPI FreeDOMTree(IXMLDOMNode* pNode)
{
	SAFE_RELEASE(pNode);
}

/**
 * @brief
 * �w��m�[�h�̎q�m�[�h�����X�g�Ɋi�[���܂��B
 * 
 * @param pNode
 * �q�m�[�h���擾���錳�ƂȂ�e�m�[�h�B
 * 
 * @param pNodeList
 * �q�m�[�h�̊i�[���IXMLDOMNodeList�̃|�C���^�̃|�C���^�B
 * 
 * @returns
 * ���X�g�Ɋi�[�����q�m�[�h����Ԃ��܂��B
 * 
 * �w��m�[�h�̎q�m�[�h�����X�g�Ɋi�[���܂��B
 * 
 * @remarks
 * ���ꂼ��̃m�[�h�I�u�W�F�N�g��GetListItem()�Ŏ擾���邱�Ƃ��ł��܂��B
 * 
 * @see
 * GetListItem()
 */
LONG WINAPI GetChildList(IXMLDOMNode* pNode, IXMLDOMNodeList** pNodeList)
{
	HRESULT hr = S_OK;
	LONG childCount = 0;

	CHECK_HR(pNode->get_childNodes(pNodeList));
	childCount = GetChildCount(*pNodeList);

Cleanup:
	return childCount;
}

/**
 * @brief
 * ���X�g�I�u�W�F�N�g�̉�����s���܂��B
 * 
 * @param pNodeList
 * IXMLDOMNodeList�̃|�C���^�B
 * 
 * GetChildList()�Ŏ擾�������X�g�I�u�W�F�N�g�̉�����s���܂��B
 * 
 * @remarks
 * GetChildList()�Ŏ擾�������X�g�I�u�W�F�N�g�͖{�֐��ŉ�����s���܂��B
 * 
 * @see
 * GetChildList()
 */
void WINAPI FreeChildList(IXMLDOMNodeList* pNodeList)
{
	SAFE_RELEASE(pNodeList);
}

/**
 * @brief
 * �w�胊�X�g�̃I�u�W�F�N�g[�m�[�h]����Ԃ��܂��B
 * 
 * @param pNodeList
 * �J�E���g�Ώۂ�IXMLDOMNodeList�|�C���^�B
 * 
 * @returns
 * �w�胊�X�g�̃I�u�W�F�N�g[�m�[�h]���B
 * 
 * �w�胊�X�g�̃I�u�W�F�N�g[�m�[�h]����Ԃ��܂��B
 */
LONG WINAPI GetChildCount(IXMLDOMNodeList* pNodeList)
{
	HRESULT hr = S_OK;
	LONG childCount = 0;
	CHECK_HR(pNodeList->get_length(&childCount));

Cleanup:
	return childCount;
}

/**
 * @brief
 * ���X�g�I�u�W�F�N�g�������index�Ɏw�肳�ꂽ�C���f�b�N�X�̃m�[�h�I�u�W�F�N�g���擾���܂��B
 * 
 * @param pNodeList
 * �ǂݎ��Ώۂ̃��X�g�I�u�W�F�N�g�B
 * 
 * @param index
 * �擾�Ώۂ̃m�[�h������킷���X�g���̃C���f�b�N�X�B
 * 
 * @param pNode
 * �擾�����m�[�h�I�u�W�F�N�g���i�[����IXMLDOMNode�̃|�C���^�̃|�C���^�B
 * 
 * @returns
 * index�Ɏw�肳���m�[�h�I�u�W�F�N�g�̎擾�ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * ���X�g�I�u�W�F�N�g�������index�Ɏw�肳�ꂽ�C���f�b�N�X�̃m�[�h�I�u�W�F�N�g���擾���܂��B
 * 
 * @remarks
 * ���X�g�̐�����GetChildList()�ōs���܂��B
 * 
 * @see
 * GetChildList()
 */
BOOL WINAPI GetListItem(IXMLDOMNodeList* pNodeList, LONG index, IXMLDOMNode** pNode)
{
	HRESULT hr = S_FALSE;
	if (pNodeList != NULL) {
		CHECK_HR(pNodeList->get_item(index, pNode));
	}

Cleanup:
	return hr==S_OK;
}

/**
 * @brief
 * �擾�������X�g���̃A�C�e��[�m�[�h�I�u�W�F�N�g]��������܂��B
 * 
 * @param pNode
 * ����Ώۂ�IXMLDOMNode�|�C���^�B
 * 
 * GetListItem()�Ŏ擾�������X�g���̃A�C�e��[�m�[�h�I�u�W�F�N�g]��������܂��B
 * 
 * @remarks
 * GetListItem()�Ŏ擾�����m�[�h�I�u�W�F�N�g�͖{�֐��ŉ�����s���܂��B
 * 
 * @see
 * GetListItem()
 */
void WINAPI FreeListItem(IXMLDOMNode* pNode)
{
	SAFE_RELEASE(pNode);
}

/**
 * @brief
 * �m�[�h�����擾���܂��B
 * 
 * @param pNode
 * �Ώۂ̃m�[�h�I�u�W�F�N�gIXMLDOMNode�̃|�C���^�B
 * 
 * @param resBuffer
 * �擾�����m�[�h�����i�[����char�o�b�t�@�B
 * 
 * @param cchLength
 * �i�[��̃o�b�t�@�̃T�C�Y�i�z�񐔁j�B
 * 
 * @returns
 * �m�[�h���̎擾�ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * �m�[�h[�^�O]�����擾���܂��B
 * 
 * @remarks
 * �m�[�h�̒l�̎擾�ɂ�GetNodeValue()�A�������̎擾�ɂ�GetAttribute()���g�p���܂��B
 * 
 * @see
 * GetNodeValue() | GetAttribute()
 */
BOOL WINAPI GetNodeName(IXMLDOMNode* pNode, PTSTR resBuffer, SIZE_T cchLength)
{
	HRESULT hr = S_OK;
	BSTR bstrName = NULL;
	CHECK_HR(pNode->get_nodeName(&bstrName));
	if (bstrName != NULL) {
		_tcscpy_s(resBuffer, cchLength, bstrName);
	}

Cleanup:
	SysFreeString(bstrName);
	return hr==S_OK;
}

/**
 * @brief
 * �m�[�h�̒l[�o�����[]���擾���܂��B
 * 
 * @param pNode
 * �Ώۂ̃m�[�h�I�u�W�F�N�gIXMLDOMNode�̃|�C���^�B
 * 
 * @param resBuffer
 * �擾�����m�[�h�̒l���i�[����char�o�b�t�@�B
 * 
 * @param cchLength
 * �i�[��̃o�b�t�@�̃T�C�Y�i�z�񐔁j�B
 * 
 * @returns
 * �m�[�h�̒l�̎擾�ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * �m�[�h�̒l[�o�����[]���擾���܂��B
 * 
 * @remarks
 * �m�[�h���̎擾�ɂ�GetNodeName()�A�������̎擾�ɂ�GetAttribute()���g�p���܂��B
 * 
 * @see
 * GetNodeName() | GetAttribute()
 */
BOOL WINAPI GetNodeValue(IXMLDOMNode* pNode, PTSTR resBuffer, SIZE_T cchLength)
{
	HRESULT hr = S_OK;
	VARIANT varNodeValue;

	VariantInit(&varNodeValue);
	CHECK_HR(pNode->get_nodeValue(&varNodeValue));
	if (resBuffer != NULL) {
		_tcscpy_s(resBuffer, cchLength, varNodeValue.bstrVal);
	}

Cleanup:
	VariantClear(&varNodeValue);
	return hr==S_OK;
}

/**
 * @brief
 * �m�[�h�̑�����[�A�g���r���[�g]���擾���܂��B
 * 
 * @param pNode
 * �Ώۂ̃m�[�h�I�u�W�F�N�gIXMLDOMNode�̃|�C���^�B
 * 
 * @param attrName
 * Description of parameter attrName.
 * 
 * @param attrValue
 * �擾�����m�[�h�̑��������i�[����char�o�b�t�@�B
 * 
 * @param cchLength
 * �i�[��̃o�b�t�@�̃T�C�Y�i�z�񐔁j�B
 * 
 * @returns
 * �m�[�h�̑������̎擾�ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * �m�[�h�̑�����[�A�g���r���[�g]���擾���܂��B
 * 
 * @remarks
 * �m�[�h���̎擾�ɂ�GetNodeName()�A�m�[�h�̒l�̎擾�ɂ�GetNodeValue()���g�p���܂��B
 * 
 * @see
 * GetNodeName() | GetNodeValue()
 */
BOOL WINAPI GetAttribute(IXMLDOMNode* pNode, PCTSTR attrName, PTSTR attrValue, SIZE_T cchLength)
{
	HRESULT hr = S_OK;
	LONG attributeCount = 0;
	IXMLDOMNamedNodeMap* pNodeMap = NULL;
	IXMLDOMNode* pAttrItem = NULL;
	BSTR bstrAttrName = NULL;
	VARIANT varAttrValue = {0};

	CHECK_HR(pNode->get_attributes(&pNodeMap));
	CHECK_HR(pNodeMap->get_length(&attributeCount));
	for (int i = 0; i < attributeCount; i++) {
		pNodeMap->get_item(i, &pAttrItem);
		if (pAttrItem != NULL) {
			VariantInit(&varAttrValue);

			pAttrItem->get_nodeName(&bstrAttrName);
			if (bstrAttrName != NULL && attrName != NULL && _tcscmp(bstrAttrName, attrName) == 0) {
				pAttrItem->get_nodeValue(&varAttrValue);
				_tcscpy_s(attrValue, cchLength, varAttrValue.bstrVal);
			}

			SysFreeString(bstrAttrName);
			VariantClear(&varAttrValue);
			SAFE_RELEASE(pAttrItem);
		}
	}

Cleanup:
	SAFE_RELEASE(pNodeMap);
	return hr==S_OK;
}

/**
 * @brief
 * Write brief comment for StoreValues here.
 * 
 * @param pNode
 * Description of parameter pNode.
 * 
 * @param attrName
 * Description of parameter attrName.
 * 
 * @returns
 * Write description of return value here.
 * 
 * @throws <exception class>
 * Description of criteria for throwing this exception.
 * 
 * Write detailed description for StoreValues here.
 * 
 * @remarks
 * Write remarks for StoreValues here.
 * 
 * @see
 * Separate items with the '|' character.
 */
map<PCTSTR, PCTSTR>* WINAPI StoreValues(IXMLDOMNode* pNode, PCTSTR attrName)
{
	map<PCTSTR, PCTSTR>* pMap = new map<PCTSTR, PCTSTR>;
	if (pMap != NULL) {
		if (!StoreValues(pNode, attrName, pMap)) {
			delete pMap;
			pMap = NULL;
		}
	}
	return pMap;
}

/**
 * @brief
 * �q�m�[�h�E�Z��m�[�h���ċA�I�Ɏ擾���Ȃ���Y���^�O�̓��e���}�b�v�Ɋi�[���܂��B
 * 
 * @param pNode
 * �ΏۂƂȂ�e�m�[�h�B
 * 
 * @param attrName
 * �}�b�v�̃L�[�ƂȂ�A�m�[�h�̃A�g���r���[�g���i�A�g���r���[�g���ɑΉ�����l���擾���A������}�b�v�̃L�[�Ƃ���j�B
 * 
 * @param keysMap
 * �i�[��̃}�b�v�B�}�b�v�͖{DLL�Ŋm�ۂ��ꂽ���̂Ƃ��܂��B
 * 
 * @returns
 * �}�b�v�ւ̊i�[�ɐ��������ꍇ�ɂ�TRUE�A���s�����ꍇ�ɂ�FALSE��Ԃ��܂��B
 * 
 * �q�m�[�h�E�Z��m�[�h���ċA�I�Ɏ擾���Ȃ���Y���^�O�݂̂����o���A���e���}�b�v�Ɋi�[���܂��B
 * �X�̃m�[�h����̃}�b�v�ւ̊i�[��Store()�ōs���܂��B
 * 
 * @remarks
 * �w�b�_�[�t�@�C���ɂ����J����Ă���StoreValues�֐��̃��[�J�[���\�b�h�ł��B
 * 
 * @see
 * StoreValues() | Store()
 */
BOOL StoreValues(IXMLDOMNode* pNode, PCTSTR attrName, map<PCTSTR, PCTSTR>* keysMap)
{
	HRESULT hr = S_OK;
	IXMLDOMNode* pChild = NULL;
	IXMLDOMNode* pSibling = NULL;

	CHECK_HR(pNode->get_firstChild(&pChild));

	while (pChild != NULL) {
		Store(pChild, attrName, keysMap);

		// �ċN�Ăяo��
		StoreValues(pChild, attrName, keysMap);
		pChild->get_nextSibling(&pSibling);
		
		SAFE_RELEASE(pChild);
		pChild = pSibling;
		pSibling = NULL;
	}
	SAFE_RELEASE(pChild);
	SAFE_RELEASE(pSibling);
	return TRUE;

Cleanup:
	SAFE_RELEASE(pChild);
	SAFE_RELEASE(pSibling);
	return FALSE;
}

/**
 * @brief
 * StoreValues()�Ŏ擾�����}�b�v��������܂��B
 * 
 * @param keysMap
 * StoreValues()�Ŏ擾�����}�b�v
 * 
 * StoreValues()�Ŏ擾�����}�b�v��������܂��B
 * �{�֐��ł̓}�b�v�Ɗi�[����Ă���L�[�A�o�����[�����S�ɉ�����܂��B
 * 
 * @remarks
 * StoreValues()�Ŋi�[�����}�b�v�͖{�֐��ŉ������悤�ɂ��Ă��������B
 * 
 * @see
 * StoreValues()
 */
void WINAPI CleanupStoredValues(map<PCTSTR, PCTSTR>* keysMap)
{
	map<PCTSTR, PCTSTR>::iterator it = keysMap->begin();
	while (it != keysMap->end()) {
		delete [] it->first;
		delete [] it->second;
		++it;
	}
	keysMap->clear();
	delete keysMap;
	keysMap = NULL;
}

/**
 * @brief
 * StoreValues()�ɂ���Ď擾���ꂽ�}�b�v�̒l��Ԃ��܂��B
 * 
 * @param pMap
 * StoreValues()�ɂ���Ď擾���ꂽ�}�b�v�B
 * 
 * @param szKey
 * �擾����l�̃L�[�B
 * 
 * @returns
 * �w��L�[�ɑΉ�����l�����݂����ꍇ�ɂ͂��̒l���A�L�[��������Ȃ��ꍇ��NULL��Ԃ��܂��B
 * 
 * StoreValues()�ɂ���Ď擾���ꂽ�}�b�v�̒l��Ԃ��܂��B
 * StoreValues()�Ŏ擾���ꂽ�}�b�v�̃A�N�Z�X�͖{�֐��ōs���Ă��������B
 * 
 * @remarks
 * StoreValues()�Ń}�b�v���擾�ł��܂��B
 * 
 * @see
 * StoreValues()
 */
PCTSTR WINAPI GetMapItem(map<PCTSTR, PCTSTR>* pMap, PCTSTR szKey)
{
	map<PCTSTR, PCTSTR>::iterator it = pMap->begin();
	while (it != pMap->end()) {
		if (!_tcsicmp(szKey, it->first)) {
			return it->second;
		}
		++it;
	}
	return NULL;
}

////////////////////////////////////////////////////////////////////////////
/**
 * @brief
 * �G���[���b�Z�[�W�_�C�A���O�𐶐����܂��B
 * 
 * @param szMessage
 * �_�C�A���O�̃��b�Z�[�W�B
 * 
 * �G���[���b�Z�[�W�_�C�A���O�𐶐����܂��B
 * 
 */
void ReportError(PCTSTR szMessage)
{
	MessageBox(NULL, szMessage, szTitle, MB_OK | MB_ICONERROR);
}

/**
 * @brief
 * �w��m�[�h�̃A�g���r���[�g�ƒl���擾���A�}�b�v�Ɋi�[���܂��B
 * 
 * @param pNode
 * �Ώۂ̃m�[�h�I�u�W�F�N�g�B
 * 
 * @param szAttrName
 * �}�b�v�̃L�[�ƂȂ�A�m�[�h�̃A�g���r���[�g���i�A�g���r���[�g���ɑΉ�����l���擾���A������}�b�v�̃L�[�Ƃ���j�B
 * 
 * @param keysMap
 * �i�[��̃}�b�v�B�}�b�v�͖{DLL�Ŋm�ۂ��ꂽ���̂Ƃ��܂��B
 * 
 * @throws <exception class>
 * Description of criteria for throwing this exception.
 * 
 * �w��m�[�h�̃A�g���r���[�g�ƒl���擾���A�q�[�v�Ɋi�[��A�}�b�v�Ɋi�[���܂��B
 * 
 * @remarks
 * �{�֐��Ŋm�ۂ����L�[�ƃo�����[�̃������̈�́ACleanupStoredValues()�ŉ������܂��B
 * 
 * @see
 * StoreValues() | CleanupStoredValues()
 */
void Store(IXMLDOMNode* pNode, PCTSTR szAttrName, map<PCTSTR, PCTSTR>* keysMap)
{
	HRESULT hr = S_OK;
	const int bufferSize = MAX_PATH;

	DOMNodeType nodeType = NODE_INVALID;
	IXMLDOMNode* pParent = NULL;
	TCHAR *szKey = NULL;
	TCHAR *szValue = NULL;

	CHECK_HR(pNode->get_nodeType(&nodeType));
	if (nodeType == NODE_TEXT) {
		CHECK_HR(pNode->get_parentNode(&pParent));

		//__try
		//{
		try {
			szKey	= new TCHAR[bufferSize];	// CleanupStoredValues()�ŉ��
			szValue	= new TCHAR[bufferSize];	// CleanupStoredValues()�ŉ��
			if (szKey == NULL || szValue == NULL) {
				if (szKey) {
					delete [] szKey;
				}
				if (szValue) {
					delete [] szValue;
				}
				ReportError(_T(MESSAGE_ERROR_MEMORY_INVALID));
				XMLParserExit(EXIT_FAILURE);
				return;
			}
		} catch (bad_alloc &ba) {
			if (szKey) {
				delete [] szKey;
			}
			if (szValue) {
				delete [] szValue;
			}
			MessageBoxA(NULL, ba.what(), "XMLParser", MB_OK | MB_ICONERROR);
			XMLParserExit(EXIT_FAILURE);
			return;
		}
		//}
		//__except (dump_exception(GetExceptionInformation())){
		//		if (szKey) {
		//			delete [] szKey;
		//		}
		//		if (szValue) {
		//			delete [] szValue;
		//		}
		//		TCHAR szError[bufferSize] = {0};
		//		_stprintf_s(szError, _countof(szError), _T("%s <error: %d>"), _T(MESSAGE_ERROR_UNKNOWN), GetExceptionCode());
		//		ReportError(szError);
		//		XMLParserExit(EXIT_FAILURE);
		//}

		ZeroMemory(szKey, sizeof(TCHAR) * bufferSize);
		ZeroMemory(szValue, sizeof(TCHAR) * bufferSize);
		GetAttribute(pParent, szAttrName, szKey, bufferSize);
		GetNodeValue(pNode, szValue, bufferSize);

		keysMap->insert(map<PCTSTR, PCTSTR>::value_type(szKey, szValue));
	}

Cleanup:
	return;
}

//LONG dump_exception(_EXCEPTION_POINTERS *ep)
//{
//	PEXCEPTION_RECORD rec = ep->ExceptionRecord;
//
//	TCHAR buf[1024];
//	_stprintf_s(buf, _countof(buf), _T("code:%x flag:%x addr:%p params:%d\n"),
//		rec->ExceptionCode,
//		rec->ExceptionFlags,
//		rec->ExceptionAddress,
//		rec->NumberParameters
//	);
//	::OutputDebugString(buf);
//
//	for (DWORD i = 0; i < rec->NumberParameters; i++){
//		_stprintf_s(buf, _countof(buf), _T("param[%d]:%x\n"),
//			i,
//			rec->ExceptionInformation[i]
//		);
//		::OutputDebugString(buf);
//	}
//
//	return EXCEPTION_EXECUTE_HANDLER;
//}

void XMLParserExit(int errorCode)
{
	exit(errorCode);
}
