#ifndef _XML_PARSER_
#define _XML_PARSER_

#pragma comment(linker, "/export:CleanupStoredValues=_CleanupStoredValues@4")
#pragma comment(linker, "/export:FreeChildList=_FreeChildList@4")
#pragma comment(linker, "/export:FreeDOMTree=_FreeDOMTree@4")
#pragma comment(linker, "/export:FreeListItem=_FreeListItem@4")
#pragma comment(linker, "/export:GetAttribute=_GetAttribute@16")
#pragma comment(linker, "/export:GetChildCount=_GetChildCount@4")
#pragma comment(linker, "/export:GetChildList=_GetChildList@8")
#pragma comment(linker, "/export:GetDOMTree=_GetDOMTree@12")
#pragma comment(linker, "/export:GetMapItem=_GetMapItem@8")
#pragma comment(linker, "/export:GetListItem=_GetListItem@12")
#pragma comment(linker, "/export:GetNodeName=_GetNodeName@12")
#pragma comment(linker, "/export:GetNodeValue=_GetNodeValue@12")
#pragma comment(linker, "/export:Initialize=_Initialize@8")
#pragma comment(linker, "/export:StoreValues=_StoreValues@8")
#pragma comment(linker, "/export:UnInitialize=_UnInitialize@4")

#include <objbase.h>	// COMで必要なヘッダーファイル
#include <msxml6.h>
#include <map>

using namespace std;

#undef DLL_EXPORT
#ifdef XMLPARSER_EXPORTS
#define DLL_EXPORT		__declspec(dllexport)
#else
#define DLL_EXPORT		__declspec(dllimport)
#endif

#ifdef __cplusplus
extern "C" {
#endif
	DLL_EXPORT BOOL WINAPI Initialize(IXMLDOMElement** pRootElement, PCTSTR szXMLuri);
	DLL_EXPORT void WINAPI UnInitialize(IXMLDOMElement* pRootElement);

	// 指定されたタグのノードを返します。メモリリークを防ぐために使用後はFreeDOMTree()を呼ぶか、Releaseしてください。
	DLL_EXPORT BOOL WINAPI GetDOMTree(IXMLDOMElement* pRootElement, PCTSTR szNodeName, IXMLDOMNode** pNode);
	DLL_EXPORT void WINAPI FreeDOMTree(IXMLDOMNode* pNode);

	// 子ノードのリストを取得し、その数を返します。メモリリークを防ぐために使用後はFreeChildList()を呼ぶか、Releaseしてください。
	DLL_EXPORT LONG WINAPI GetChildList(IXMLDOMNode* pNode, IXMLDOMNodeList** pNodeList);
	DLL_EXPORT void WINAPI FreeChildList(IXMLDOMNodeList* pNodeList);
	// 子ノードの数を返します。
	DLL_EXPORT LONG WINAPI GetChildCount(IXMLDOMNodeList* pNodeList);
	// 子ノードのリストから、indexに示されるノードを返します。メモリリークを防ぐために使用後はFreeListItem()を呼ぶか、Releaseしてください。
	DLL_EXPORT BOOL WINAPI GetListItem(IXMLDOMNodeList* pNodeList, LONG index, IXMLDOMNode** pNode);
	DLL_EXPORT void WINAPI FreeListItem(IXMLDOMNode* pNode);

	// ノードの名前を返します。
	// lengthにはTCHAR配列の個数を指定してください。
	DLL_EXPORT BOOL WINAPI GetNodeName(IXMLDOMNode* pNode, PTSTR resBuffer, SIZE_T cchLength);
	// ノードの値を返します。
	DLL_EXPORT BOOL WINAPI GetNodeValue(IXMLDOMNode* pNode, PTSTR resBuffer, SIZE_T cchLength);

	// ノードに含まれるアトリビュートの中からattrNameに該当するキーに対応する値を返します。
	DLL_EXPORT BOOL WINAPI GetAttribute(IXMLDOMNode* pNode, PCTSTR attrName, PTSTR attrValue, SIZE_T cchLength);

	
	// <node name="pika">moko</node>
	// -> attrName = TEXT("name");　とすると keysMapにはkey=pika, value=mokoの組み合わせで登録します
	// StoreValues使用後は、必ずCleanupStoredValuesを呼んでください。
	DLL_EXPORT map<PCTSTR, PCTSTR>* WINAPI StoreValues(IXMLDOMNode* pNode, PCTSTR attrName);
	DLL_EXPORT void WINAPI CleanupStoredValues(map<PCTSTR, PCTSTR>* keysMap);

	// StoreValuesで作成したMapへのアクセサです。
	DLL_EXPORT PCTSTR WINAPI GetMapItem(map<PCTSTR, PCTSTR>* pMap, PCTSTR szKey);

	void XMLParserExit(int errorCode);

#ifdef __cplusplus
}
#endif

#endif /* _XML_PARSER_ */