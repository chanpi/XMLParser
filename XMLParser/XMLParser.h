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

#include <objbase.h>	// COM�ŕK�v�ȃw�b�_�[�t�@�C��
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

	// �w�肳�ꂽ�^�O�̃m�[�h��Ԃ��܂��B���������[�N��h�����߂Ɏg�p���FreeDOMTree()���ĂԂ��ARelease���Ă��������B
	DLL_EXPORT BOOL WINAPI GetDOMTree(IXMLDOMElement* pRootElement, PCTSTR szNodeName, IXMLDOMNode** pNode);
	DLL_EXPORT void WINAPI FreeDOMTree(IXMLDOMNode* pNode);

	// �q�m�[�h�̃��X�g���擾���A���̐���Ԃ��܂��B���������[�N��h�����߂Ɏg�p���FreeChildList()���ĂԂ��ARelease���Ă��������B
	DLL_EXPORT LONG WINAPI GetChildList(IXMLDOMNode* pNode, IXMLDOMNodeList** pNodeList);
	DLL_EXPORT void WINAPI FreeChildList(IXMLDOMNodeList* pNodeList);
	// �q�m�[�h�̐���Ԃ��܂��B
	DLL_EXPORT LONG WINAPI GetChildCount(IXMLDOMNodeList* pNodeList);
	// �q�m�[�h�̃��X�g����Aindex�Ɏ������m�[�h��Ԃ��܂��B���������[�N��h�����߂Ɏg�p���FreeListItem()���ĂԂ��ARelease���Ă��������B
	DLL_EXPORT BOOL WINAPI GetListItem(IXMLDOMNodeList* pNodeList, LONG index, IXMLDOMNode** pNode);
	DLL_EXPORT void WINAPI FreeListItem(IXMLDOMNode* pNode);

	// �m�[�h�̖��O��Ԃ��܂��B
	// length�ɂ�TCHAR�z��̌����w�肵�Ă��������B
	DLL_EXPORT BOOL WINAPI GetNodeName(IXMLDOMNode* pNode, PTSTR resBuffer, SIZE_T cchLength);
	// �m�[�h�̒l��Ԃ��܂��B
	DLL_EXPORT BOOL WINAPI GetNodeValue(IXMLDOMNode* pNode, PTSTR resBuffer, SIZE_T cchLength);

	// �m�[�h�Ɋ܂܂��A�g���r���[�g�̒�����attrName�ɊY������L�[�ɑΉ�����l��Ԃ��܂��B
	DLL_EXPORT BOOL WINAPI GetAttribute(IXMLDOMNode* pNode, PCTSTR attrName, PTSTR attrValue, SIZE_T cchLength);

	
	// <node name="pika">moko</node>
	// -> attrName = TEXT("name");�@�Ƃ���� keysMap�ɂ�key=pika, value=moko�̑g�ݍ��킹�œo�^���܂�
	// StoreValues�g�p��́A�K��CleanupStoredValues���Ă�ł��������B
	DLL_EXPORT map<PCTSTR, PCTSTR>* WINAPI StoreValues(IXMLDOMNode* pNode, PCTSTR attrName);
	DLL_EXPORT void WINAPI CleanupStoredValues(map<PCTSTR, PCTSTR>* keysMap);

	// StoreValues�ō쐬����Map�ւ̃A�N�Z�T�ł��B
	DLL_EXPORT PCTSTR WINAPI GetMapItem(map<PCTSTR, PCTSTR>* pMap, PCTSTR szKey);

	void XMLParserExit(int errorCode);

#ifdef __cplusplus
}
#endif

#endif /* _XML_PARSER_ */