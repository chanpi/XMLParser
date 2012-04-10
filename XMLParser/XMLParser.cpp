/**
 * XMLParser.cpp
 * 
 * Copyright (c) 2008 by <your name/ organization here>
 */
// XMLParser.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
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
 * COMの初期化とXMLのロードを行います。
 * 
 * @param pRootElement
 * IXMLDOMElementを格納するポインタのポインタ。
 * 
 * @param szXMLuri
 * ロードするXMLファイルのパス。
 * 
 * @returns
 * 初期化およびロードが成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * COMの初期化とXMLのロードを行います。
 * 
 * @remarks
 * 関数が成功した場合には、プログラム終了前にUnInitialize()を呼び出してください。
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
 * COMの終了処理とIXMLDOMElementの解放を行います。
 * 
 * @param pRootElement
 * Initializeによって格納されたIXMLDOMElementポインタ。
 * 
 * COMの終了処理とIXMLDOMElementの解放を行います。
 * 
 * @remarks
 * Initializeによって関数が成功した場合には、プログラム終了前に本関数を呼び出してください。
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

// ↓メモリリークする可能性があるので使用禁止。
// SysFreeStringを呼ぶ必要がある。
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
 * DOMオブジェクトの生成との初期化を行います。
 * 
 * @param ppDoc
 * IXMLDOMDocumentのポインタのポインタ。
 * 
 * @returns
 * 処理結果を返します。
 * 
 * DOMオブジェクトの生成との初期化を行います。
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
 * XMLをロードし、IXMLDOMElementオブジェクトに格納します。
 * 
 * @param pRootElement
 * XML中のルートエレメントを格納するオブジェクトのポインタのポインタ。
 * 
 * @param szXMLuri
 * ロードするXMLファイルのパス。
 * 
 * @returns
 * ロードおよびオブジェクト化に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * XMLをロードし、IXMLDOMElementオブジェクトに格納します。
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

		// ルートエレメント取得
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
 * 指定したルートエレメントからノードオブジェクトを取得します。
 * 
 * @param pRootElement
 * ロードしたXMLが格納されたルートエレメント。
 * 
 * @param szNodeName
 * ノードオブジェクト対象のノード[タグ]名。
 * 
 * @param pNode
 * 格納するIXMLDOMNodeのポインタのポインタ。
 * 
 * @returns
 * ノードオブジェクトの取得に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * 指定したルートエレメントから、指定ノードのノードオブジェクトを取得します。
 * 
 * @remarks
 * 使用後はFreeDOMTree()を呼び出してください。
 * 
 * @see
 * FreeDOMTree()
 */
BOOL WINAPI GetDOMTree(IXMLDOMElement* pRootElement, PCTSTR szNodeName, IXMLDOMNode** pNode)
{
	HRESULT hr = S_OK;
	if (pRootElement == NULL) {
		//ReportError(TEXT("先にXMLの読み込みを行ってください。"));
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
 * ノードオブジェクトの解放を行います。
 * 
 * @param pNode
 * IXMLDOMNodeのポインタ。
 * 
 * GetDOMTree()で取得したノードオブジェクトの解放を行います。
 * 
 * @remarks
 * GetDOMTree()で取得したノードオブジェクトは本関数で解放を行います。
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
 * 指定ノードの子ノードをリストに格納します。
 * 
 * @param pNode
 * 子ノードを取得する元となる親ノード。
 * 
 * @param pNodeList
 * 子ノードの格納先のIXMLDOMNodeListのポインタのポインタ。
 * 
 * @returns
 * リストに格納した子ノード数を返します。
 * 
 * 指定ノードの子ノードをリストに格納します。
 * 
 * @remarks
 * それぞれのノードオブジェクトはGetListItem()で取得することができます。
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
 * リストオブジェクトの解放を行います。
 * 
 * @param pNodeList
 * IXMLDOMNodeListのポインタ。
 * 
 * GetChildList()で取得したリストオブジェクトの解放を行います。
 * 
 * @remarks
 * GetChildList()で取得したリストオブジェクトは本関数で解放を行います。
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
 * 指定リストのオブジェクト[ノード]数を返します。
 * 
 * @param pNodeList
 * カウント対象のIXMLDOMNodeListポインタ。
 * 
 * @returns
 * 指定リストのオブジェクト[ノード]数。
 * 
 * 指定リストのオブジェクト[ノード]数を返します。
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
 * リストオブジェクトから引数indexに指定されたインデックスのノードオブジェクトを取得します。
 * 
 * @param pNodeList
 * 読み取り対象のリストオブジェクト。
 * 
 * @param index
 * 取得対象のノードをあらわすリスト中のインデックス。
 * 
 * @param pNode
 * 取得したノードオブジェクトを格納するIXMLDOMNodeのポインタのポインタ。
 * 
 * @returns
 * indexに指定されるノードオブジェクトの取得に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * リストオブジェクトから引数indexに指定されたインデックスのノードオブジェクトを取得します。
 * 
 * @remarks
 * リストの生成はGetChildList()で行います。
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
 * 取得したリスト中のアイテム[ノードオブジェクト]を解放します。
 * 
 * @param pNode
 * 解放対象のIXMLDOMNodeポインタ。
 * 
 * GetListItem()で取得したリスト中のアイテム[ノードオブジェクト]を解放します。
 * 
 * @remarks
 * GetListItem()で取得したノードオブジェクトは本関数で解放を行います。
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
 * ノード名を取得します。
 * 
 * @param pNode
 * 対象のノードオブジェクトIXMLDOMNodeのポインタ。
 * 
 * @param resBuffer
 * 取得したノード名を格納するcharバッファ。
 * 
 * @param cchLength
 * 格納先のバッファのサイズ（配列数）。
 * 
 * @returns
 * ノード名の取得に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * ノード[タグ]名を取得します。
 * 
 * @remarks
 * ノードの値の取得にはGetNodeValue()、属性名の取得にはGetAttribute()を使用します。
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
 * ノードの値[バリュー]を取得します。
 * 
 * @param pNode
 * 対象のノードオブジェクトIXMLDOMNodeのポインタ。
 * 
 * @param resBuffer
 * 取得したノードの値を格納するcharバッファ。
 * 
 * @param cchLength
 * 格納先のバッファのサイズ（配列数）。
 * 
 * @returns
 * ノードの値の取得に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * ノードの値[バリュー]を取得します。
 * 
 * @remarks
 * ノード名の取得にはGetNodeName()、属性名の取得にはGetAttribute()を使用します。
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
 * ノードの属性名[アトリビュート]を取得します。
 * 
 * @param pNode
 * 対象のノードオブジェクトIXMLDOMNodeのポインタ。
 * 
 * @param attrName
 * Description of parameter attrName.
 * 
 * @param attrValue
 * 取得したノードの属性名を格納するcharバッファ。
 * 
 * @param cchLength
 * 格納先のバッファのサイズ（配列数）。
 * 
 * @returns
 * ノードの属性名の取得に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * ノードの属性名[アトリビュート]を取得します。
 * 
 * @remarks
 * ノード名の取得にはGetNodeName()、ノードの値の取得にはGetNodeValue()を使用します。
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
 * 子ノード・兄弟ノードを再帰的に取得しながら該当タグの内容をマップに格納します。
 * 
 * @param pNode
 * 対象となる親ノード。
 * 
 * @param attrName
 * マップのキーとなる、ノードのアトリビュート名（アトリビュート名に対応する値を取得し、これをマップのキーとする）。
 * 
 * @param keysMap
 * 格納先のマップ。マップは本DLLで確保されたものとします。
 * 
 * @returns
 * マップへの格納に成功した場合にはTRUE、失敗した場合にはFALSEを返します。
 * 
 * 子ノード・兄弟ノードを再帰的に取得しながら該当タグのみを検出し、内容をマップに格納します。
 * 個々のノードからのマップへの格納はStore()で行います。
 * 
 * @remarks
 * ヘッダーファイルにより公開されているStoreValues関数のワーカーメソッドです。
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

		// 再起呼び出し
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
 * StoreValues()で取得したマップを解放します。
 * 
 * @param keysMap
 * StoreValues()で取得したマップ
 * 
 * StoreValues()で取得したマップを解放します。
 * 本関数ではマップと格納されているキー、バリューを安全に解放します。
 * 
 * @remarks
 * StoreValues()で格納したマップは本関数で解放するようにしてください。
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
 * StoreValues()によって取得されたマップの値を返します。
 * 
 * @param pMap
 * StoreValues()によって取得されたマップ。
 * 
 * @param szKey
 * 取得する値のキー。
 * 
 * @returns
 * 指定キーに対応する値が存在した場合にはその値を、キーが見つからない場合はNULLを返します。
 * 
 * StoreValues()によって取得されたマップの値を返します。
 * StoreValues()で取得されたマップのアクセスは本関数で行ってください。
 * 
 * @remarks
 * StoreValues()でマップを取得できます。
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
 * エラーメッセージダイアログを生成します。
 * 
 * @param szMessage
 * ダイアログのメッセージ。
 * 
 * エラーメッセージダイアログを生成します。
 * 
 */
void ReportError(PCTSTR szMessage)
{
	MessageBox(NULL, szMessage, szTitle, MB_OK | MB_ICONERROR);
}

/**
 * @brief
 * 指定ノードのアトリビュートと値を取得し、マップに格納します。
 * 
 * @param pNode
 * 対象のノードオブジェクト。
 * 
 * @param szAttrName
 * マップのキーとなる、ノードのアトリビュート名（アトリビュート名に対応する値を取得し、これをマップのキーとする）。
 * 
 * @param keysMap
 * 格納先のマップ。マップは本DLLで確保されたものとします。
 * 
 * @throws <exception class>
 * Description of criteria for throwing this exception.
 * 
 * 指定ノードのアトリビュートと値を取得し、ヒープに格納後、マップに格納します。
 * 
 * @remarks
 * 本関数で確保したキーとバリューのメモリ領域は、CleanupStoredValues()で解放されます。
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
			szKey	= new TCHAR[bufferSize];	// CleanupStoredValues()で解放
			szValue	= new TCHAR[bufferSize];	// CleanupStoredValues()で解放
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
