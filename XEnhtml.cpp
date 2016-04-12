#include "stdafx.h"
#include "XEncoding.h"
#include "XMemStream.h"

#define	ERROR_TOKEN		258
#define	SPACETEXT		259
#define	NBSPTEXT		260
#define	PLANTEXT		261
#define	TAG_ID			262
#define	TAG_STR			263
#define	TAG_START_ID	264
#define	TAG_CLOSE_ID	265
#define STYLE_BEGIN		266
#define STYLE_END		267
#define STYLE_ID		268
#define STYLE_VALUE		269

#define TAG_MODE		1
#define STYLE_MODE		2

#define TAG_HTML	0
#define TAG_TEXT	1
#define TAG_WAP		2

#define TAG_IMG		1
#define TAG_PRE		2
#define TAG_P		3

struct _tagName{
	const WCHAR*	name;
	int	length;
	unsigned int   id;	// TAG_IMG, TAG_PRE，用于快速判断特殊 tag
	unsigned int   mode:2;  // 0：不需封口； 1: 删除； 2：需要封口
	unsigned int   before:1;	// 0: 不换行；1: tag 前换行
	unsigned int   inside:1;	// 0: 不换行；1: tag 内换行
	unsigned int   after:1;		// 0: 不换行；1: tag 后换行
	unsigned int   isWAP:1;
	const char*		modes;		// 1: 删除
	const WCHAR*	baseattr;	// 必须 attr，不包含则删除 tag
};

static _tagName s_tags[] = 
{	//name		length id	  mode b  i  a  w  mode
	{L"a", 			1, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000", L"href"},
	{L"acronym", 	7, 0,		2, 0, 0, 0, 0, "001000003200002000002000000000001000022000000000000210200000"},
	{L"address", 	7, 0,		2, 1, 1, 1, 0, "200100023200022002222000000000001000022223200000022212222002"},
	{L"b", 			1, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"bdo", 		3, 0,		2, 1, 0, 1, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"big", 		3, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"blockquote",10, 0,		2, 1, 0, 1, 0, "200200003200022002222000000000001000022223200000022212222022"},
	{L"br",			2, 0,		0, 0, 1, 0, 0, "000000000000000000000000000000001000000000000000000010000000"},
	{L"caption", 	7, 0,		2, 1, 1, 1, 0, "111111113111100111110101111111011011100111011111100010001111"},
	{L"center", 	6, 0,		2, 1, 1, 1, 0, "000000003200002000002000000000001000022003000000000210200000"},
	{L"cite", 		4, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"code", 		4, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"colgroup", 	8, 0,		2, 1, 1, 1, 0, "111111113011110111110101111111011011100111011111100010001111"},
	{L"dd", 		2, 0,		2, 1, 0, 1, 0, "200200023200000002222000000000001000022222200000022212222022"},
	{L"del", 		3, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"dfn", 		3, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"dir", 		3, 0,		2, 1, 0, 1, 0, "200200023200022000222000000000001000022223200000022212222022"},
	{L"div", 		3, 0,		2, 1, 0, 1, 0, "000200023200002002022000000000001000022223200000022212222022"},
	{L"dl", 		2, 0,		2, 1, 1, 1, 0, "200200023200022002202000000000001000022223200000022212222022"},
	{L"dt", 		2, 0,		2, 1, 0, 1, 0, "200200023200002002220000000000001000022222200000022212222022"},
	{L"em", 		2, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"embed", 		5, 0,		0, 0, 0, 0, 0, "000000003000000000000000000000001000000000000000000010000000", L"src"},
	{L"font", 		4, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"h1", 		2, 0,		2, 1, 0, 1, 0, "000000003200002000002000122222001000022003000000000210200000"},
	{L"h2", 		2, 0,		2, 1, 0, 1, 0, "000000003200002000002000212222001000022003000000000210200000"},
	{L"h3", 		2, 0,		2, 1, 0, 1, 0, "000000003200002000002000221222001000022003000000000210200000"},
	{L"h4", 		2, 0,		2, 1, 0, 1, 0, "000000003200002000002000222122001000022003000000000210200000"},
	{L"h5", 		2, 0,		2, 1, 0, 1, 0, "000000003200002000002000222212001000022003000000000210200000"},
	{L"h6", 		2, 0,		2, 1, 0, 1, 0, "000000003200002000002000222221001000022003000000000210200000"},
	{L"hr", 		2, 0,		0, 1, 1, 0, 0, "000000003000000000000000000000001000000000000000000010000000"},
	{L"i", 			1, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"iframe",		6, 0,		2, 0, 0, 0, 0, "000000000000000000000000000000001000000000000000000010000000", L"src"},
	{L"img", 		3, TAG_IMG, 0, 0, 0, 0, 1, "000000003000000000000000000000001000000000000000000010000000", L"src"},
	{L"ins", 		3, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"kbd", 		3, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"label", 		5, 0,		2, 1, 0, 1, 0, "000000003200002000002000000000001000122000000000000210200000"},
	{L"li", 		2, 0,		0, 1, 0, 0, 0, "200200023200002002222000000000001000012222200000022212222022"},
	{L"marquee", 	7, 0,		2, 1, 1, 1, 0, "300300033230000003330000333333001000021333300000022212222033"},
	{L"menu", 		4, 0,		2, 1, 1, 1, 0, "200200023200022002222000000000001000022123200000022212222022"},
	{L"ol", 		2, 0,		2, 1, 1, 1, 0, "200200023200022002222000000000001000022213200000022212222022"},
	{L"p", 			1, TAG_P,	2, 1, 0, 1, 1, "200200023200022002222000000000001000022221200000022212222022"},
	{L"pre", 		3, TAG_PRE, 2, 1, 0, 1, 0, "200200023200002002222000000000001000022222100000022212222022"},
	{L"q", 			1, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"small", 		5, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"span", 		4, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200000"},
	{L"strong", 	6, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"sub", 		3, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000001000210200000"},
	{L"sup", 		3, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000100210200000"},
	{L"table", 		5, 0,		2, 1, 1, 1, 0, "300300033000020003330000000000001000022333300000000012020033"},
	{L"tbody", 		5, 0,		2, 1, 1, 1, 0, "111111113011100111110101111111011011100111011111121010001111"},
	{L"td", 		2, 0,		2, 1, 0, 1, 0, "111111113011100111110101111111011011100111011111100110002111"},
	{L"textarea", 	8, 0,		2, 1, 0, 1, 0, "222222222222222222222222222222221222222222222222222212222222"},
	{L"tfoot", 		5, 0,		2, 1, 1, 1, 0, "111111113011100111110101111111011011100111011111100011001111"},
	{L"th", 		2, 0,		2, 1, 1, 1, 0, "111111113011100111110101111111011011100111011111100010100111"},
	{L"thead", 		5, 0,		2, 1, 1, 1, 0, "111111113011100111110101111111011011100111011111100010011111"},
	{L"tr", 		2, 0,		2, 1, 1, 1, 0, "111111113011100111110101111111011011100111011111122012021111"},
	{L"tt", 		2, 0,		2, 0, 0, 0, 0, "000000003200002000002000000000001000022000000000000210200100"},
	{L"u", 			1, 0,		2, 0, 0, 0, 1, "000000003200002000002000000000001000022000000000000210200000"},
	{L"ul", 		2, 0,		2, 1, 1, 1, 0, "200200023200022002222000000000001000022223200000022212222021"}
};

static const WCHAR* s_html_attrs[] =
{
	L"align", L"autostart", L"background", L"bgcolor", L"border", L"bordercolor", L"bordercolordark", L"bordercolorlight",
	L"cellpadding", L"cellspacing", L"clear", L"color", L"cols", L"colspan", L"dir", L"direction", L"disabled",
	L"face", L"frame", L"frameborder", L"galleryimg", L"headers", L"height", L"href", L"hspace", L"loop", L"lowsrc", L"marginheight", 
	L"marginwidth", L"rows", L"rowspan", L"rules", L"scope", L"scrollamount", L"scrolldelay", L"scrolling", L"showcontrols",
	L"size", L"span", L"src", L"start", L"summary", L"target", L"text", L"truespeed", L"type", L"units", L"unselectable",
	L"urn", L"valign", L"value", L"volume", L"vspace", L"width", L"wrap"
};

static const WCHAR* s_wap_attrs[] =
{
	L"align", L"height", L"href", L"hspace", L"src", L"vspace", L"width"
};

static const WCHAR* s_styles[] =
{
	L"background", L"background-attachment", L"background-color", L"background-image", L"background-position",
	L"background-position-x", L"background-position-y", L"background-repeat", L"border", L"border-bottom",
	L"border-bottom-color", L"border-bottom-style", L"border-bottom-width", L"border-collapse", L"border-color",
	L"border-left", L"border-left-color", L"border-left-style", L"border-left-width", L"border-right",
	L"border-right-color", L"border-right-style", L"border-right-width", L"border-style", L"border-top", L"border-top-color", 
	L"border-top-style", L"border-top-width", L"border-width", L"clear", L"color", L"direction", L"float", L"font",
	L"font-family", L"font-size", L"font-style", L"font-variant", L"font-weight", L"height", L"layout-flow", L"layout-grid",
	L"layout-grid-char", L"layout-grid-line", L"layout-grid-mode", L"layout-grid-type", L"letter-spacing", L"line-break",
	L"line-height", L"margin", L"margin-bottom", L"margin-left", L"margin-right", L"margin-top", L"min-height", L"padding",
	L"padding-bottom", L"padding-left", L"padding-right", L"padding-top", L"ruby-align", L"ruby-overhang", L"ruby-position",
	L"scrollbar-3dlight-color", L"scrollbar-arrow-color", L"scrollbar-base-color", L"scrollbar-darkshadow-color",
	L"scrollbar-face-color", L"scrollbar-highlight-color", L"scrollbar-shadow-color", L"scrollbar-track-color",
	L"table-layout", L"text-align", L"text-align-last", L"text-autospace", L"text-decoration", L"text-indent", L"text-justify",
	L"text-kashida-space", L"text-overflow", L"text-transform", L"text-underline-position", L"unicode-bidi",
	L"vertical-align", L"white-space", L"width", L"word-break", L"word-spacing", L"word-wrap"
};

static int str_compare( const WCHAR **arg1, const WCHAR **arg2 )
{
	return wcscmp( *arg1, *arg2 );
}

class CXHtmlLex
{
public:
	CXHtmlLex(BSTR str, int m)
	{
		bSpace = 0;
		bNewLine = 0;
		bSkipSpace = 1;
		nPRE = 0;
		nIndent = 0;
		nLinePos = 0;
		nTextPos = 0;
		mode = 0;
		bTagValue = FALSE;
		nTagMode = m;

		buf_text = str;
		buf_leng = ::SysStringLen(str);
	}

	void fmtHTML()
	{
		int n;
		const WCHAR* ptr;
		_tagName* pTag = NULL;
		const WCHAR** psttrs;
		int attr_count;

		if(nTagMode == TAG_WAP)
		{
			psttrs = s_wap_attrs;
			attr_count = sizeof(s_wap_attrs) / sizeof(s_wap_attrs[0]);
		}else
		{
			psttrs = s_html_attrs;
			attr_count = sizeof(s_html_attrs) / sizeof(s_html_attrs[0]);
		}

		n = fht_lex();
		while(n)
		{
			switch(n)
			{
			case '<':
				putSplit();
				m_Stream.Write(L"&lt;", 8, NULL);
				nLinePos += 4;
				nTextPos += 1;

				n = fht_lex();
				break;
			case '>':
				putSplit();
				m_Stream.Write(L"&gt;", 8, NULL);
				nLinePos += 4;
				nTextPos += 1;

				n = fht_lex();
				break;
			case NBSPTEXT:
			case SPACETEXT: 
				if(nPRE)
				{
					m_Stream.Write(fht_text, fht_leng * 2, NULL);
					nLinePos += fht_leng;
					nTextPos += fht_leng;
				}else bSpace = 1;

				n = fht_lex();
				break;

			case PLANTEXT:
				putSplit();
				m_Stream.Write(fht_text, fht_leng * 2, NULL);
				nLinePos += fht_leng;
				nTextPos += fht_leng;

				n = fht_lex();
				break;

			case TAG_START_ID:
				{
					CXString strTemp(fht_text + 1, fht_leng - 1);
					strTemp.MakeLower();
					ptr = (const WCHAR*)strTemp;
					pTag = (_tagName*)bsearch( (const WCHAR *) &ptr, (const WCHAR *)s_tags, sizeof(s_tags) / sizeof(s_tags[0]),
										sizeof( s_tags[0] ), (int (*)(const void*, const void*))str_compare );
				}

				if(!pTag)
					do{n = fht_lex();}while(n != 0 && n != '>');
				else
				{
					int nTag = int(pTag - s_tags);
					char bDel = '1';
					int i;

					if(arrayTagStack.GetCount() == 0)
						bDel = pTag->modes[0];		// 不能作为顶层的 tag
					else
					{
						for(i = (int)arrayTagStack.GetCount() - 1; i >= 0 && arrayTagStack[i] != nTag && bDel == '1'; i --)
							bDel = pTag->modes[arrayTagStack[i] + 1];	// 检查此 tag 是否与外层 tag 冲突
						if(i >= 0 && arrayTagStack[i] == nTag)
							bDel = '0';
					}

					for(i = (int)arrayTagStack.GetCount() - 1; i >= 0 && pTag->modes[arrayTagStack[i] + 1] != '2'; i --)
						if(pTag->modes[arrayTagStack[i] + 1] == '1')
							CloseTag(i);

					CRBMap<CXString, CXString> mapAttr;
					CRBMap<CXString, CXString> mapStyle;

					n = fht_lex();
					while(1)
					{
						if(n == 0 || n == '>')break;

						if(n == STYLE_BEGIN)
						{
							CXString strStyle;

							n = fht_lex();
							while(1)
							{
								if(n == 0 || n == STYLE_END)break;

								if(n != STYLE_ID)
								{
									n = fht_lex();
									continue;
								}

								CXString strTemp(fht_text, fht_leng);
								strTemp.MakeLower();

								n = fht_lex();
								if(n != ':')continue;

								bTagValue = TRUE;
								n = fht_lex();
								bTagValue = FALSE;
								if(n != STYLE_VALUE)continue;

								ptr = strTemp;
								if(!bsearch( (const WCHAR *) &ptr, (const WCHAR *)s_styles, sizeof(s_styles) / sizeof(const WCHAR *),
										sizeof( const WCHAR * ), (int (*)(const void*, const void*))str_compare ))
								{
									n = fht_lex();
									continue;
								}

								if((!strTemp.Compare(L"width") || !strTemp.Compare(L"height") || !strTemp.Compare(L"font-size")) && _wtol(fht_text) < 16)
								{
									n = fht_lex();
									continue;
								}

								if(!_wcsnicmp(fht_text, L"url(", 4))
								{
									ptr = fht_text + 4;
									if(*ptr == '\"' || *ptr == '\'')ptr ++;

									if(_wcsnicmp(ptr, L"javascript:", 11) && _wcsnicmp(ptr, L"vbscript:", 9))
										mapStyle.SetAt(strTemp, CXString(fht_text, fht_leng));
								}else mapStyle.SetAt(strTemp, CXString(fht_text, fht_leng));
							}

							if(n != '>')n = fht_lex();
							continue;
						}

						if(n != TAG_ID)
						{
							n = fht_lex();
							continue;
						}

						CXString strTemp(fht_text, fht_leng);
						strTemp.MakeLower();

						n = fht_lex();
						if(n != '=')continue;

						bTagValue = TRUE;
						n = fht_lex();
						bTagValue = FALSE;
						if(!fht_leng)continue;
						if(n != TAG_ID && n != TAG_STR)continue;

						ptr = strTemp;
						if(!bsearch( (const WCHAR *) &ptr, (const WCHAR *)psttrs, attr_count,
								sizeof( const WCHAR * ), (int (*)(const void*, const void*))str_compare ))
						{
							n = fht_lex();
							continue;
						}

						CXString strTemp1(fht_text, fht_leng);

						if(!_wcsnicmp(strTemp1, L"javascript:", 11) || !_wcsnicmp(strTemp1, L"vbscript:", 9))
						{
							n = fht_lex();
							continue;
						}

						if((!strTemp.Compare(L"width") || !strTemp.Compare(L"height")) && _wtol(strTemp1) < 16)
						{
							n = fht_lex();
							continue;
						}

						strTemp1.Replace(L"\"", L"%22");
						strTemp1.Replace(L"\'", L"%27");
						strTemp1.Replace(L"<", L"%3C");
						strTemp1.Replace(L">", L"%3E");

						strTemp1 = '\"' + strTemp1 + '\"';

						mapAttr.SetAt(strTemp, strTemp1);

						n = fht_lex();
					}

					if(pTag->baseattr && !mapAttr.Lookup(pTag->baseattr))
					{
						n = fht_lex();
						break;
					}

					if(!nPRE && pTag->before)bNewLine = 1;
					if(nTagMode == TAG_WAP)
					{
						if(nTextPos && pTag->id == TAG_IMG)
							bNewLine = 1;

						if(pTag->id == TAG_P)
							nTextPos = 0;
					}

					putSplit();

					if(nTagMode == TAG_HTML || (nTagMode == TAG_WAP && pTag->isWAP))
					{
						POSITION pos;
						CRBMap<CXString, CXString>::CPair* pPair;

						m_Stream.Write(L"<", 2, NULL);
						m_Stream.Write(pTag->name, pTag->length * 2, NULL);
						nLinePos += pTag->length + 1;
						if(pTag->id == TAG_PRE)nPRE ++;

						pos = mapAttr.GetHeadPosition();
						while(pPair = (CRBMap<CXString, CXString>::CPair*)pos)
						{
							putSpace();

							m_Stream.Write(pPair->m_key, pPair->m_key.GetLength() * 2, NULL);
							m_Stream.Write(L"=", 2, NULL);
							m_Stream.Write(pPair->m_value, pPair->m_value.GetLength() * 2, NULL);
							nLinePos += pPair->m_key.GetLength() + pPair->m_value.GetLength() + 1;

							mapAttr.GetNext(pos);
						}

						if(nTagMode != TAG_WAP)
						{
							pos = mapStyle.GetHeadPosition();
							if(pos)
							{
								putSpace();
								m_Stream.Write(L"style=\"", 14, NULL);
								nLinePos += 7;
								while(pPair = (CRBMap<CXString, CXString>::CPair*)pos)
								{
									m_Stream.Write(pPair->m_key, pPair->m_key.GetLength() * 2, NULL);
									m_Stream.Write(L":", 2, NULL);
									m_Stream.Write(pPair->m_value, pPair->m_value.GetLength() * 2, NULL);
									m_Stream.Write(L";", 2, NULL);
									nLinePos += pPair->m_key.GetLength() + pPair->m_value.GetLength() + 4;

									mapStyle.GetNext(pos);
								}
								m_Stream.Write(L"\"", 2, NULL);
								nLinePos += 1;
							}
						}

						if(pTag->mode == 0)
						{
							m_Stream.Write(L" /", 4, NULL);
							nLinePos += 2;
						}
						m_Stream.Write(L">", 2, NULL);
						nLinePos ++;
					}

					if(!nPRE && pTag->inside)
					{
						bNewLine = 1;
						if(pTag->after && pTag->before && (nTagMode == TAG_HTML))
							nIndent ++;
					}
					if(!nPRE && !pTag->inside && pTag->before)bSkipSpace = 1;

					if(pTag->mode > 1)arrayTagStack.Add(nTag);
					else if(!nPRE && pTag->after)
						bNewLine = 1;

					if(nTagMode == TAG_WAP && pTag->id == TAG_IMG)
					{
						bNewLine = 1;
						nTextPos = 1;
					}
				}

				n = fht_lex();

				break;

			case TAG_CLOSE_ID:
				{
					CXString strTemp(fht_text + 2, fht_leng - 2);
					strTemp.MakeLower();
					ptr = (const WCHAR*)strTemp;
					pTag = (_tagName*)bsearch( (const WCHAR *) &ptr, (const WCHAR *)s_tags, sizeof(s_tags) / sizeof(s_tags[0]),
							sizeof( s_tags[0] ), (int (*)(const void*, const void*))str_compare );
				}

				if(pTag && pTag->mode > 1)
				{
					int nTag = int(pTag - s_tags);
					int i;
					char bDel = pTag->modes[0];

					i = (int)arrayTagStack.GetCount() - 1;
					if(bDel == '1')
						while(i >= 0 && arrayTagStack[i] != nTag && pTag->modes[arrayTagStack[i] + 1] == '1')i --;
					else
						while(i >= 0 && arrayTagStack[i] != nTag && s_tags[arrayTagStack[i]].modes[nTag + 1] != '3')i --;

					if(i < 0 || arrayTagStack[i] != nTag)
					{
						n = fht_lex();
						break;
					}

					CloseTag(i);
				}

				n = fht_lex();
				break;
			default:
				n = fht_lex();
				break;
			};
		}

		int i = (int)arrayTagStack.GetCount() - 1;
		while(i >= 0)
		{
			CloseTag(i);
			i --;
		}
	}

	void GetString(CXString& str)
	{
		int nSize = (int)m_Stream.GetLength();
		m_Stream.SeekToBegin();
		m_Stream.Read(str.GetBuffer(nSize / 2), nSize);
		str.ReleaseBuffer(nSize / 2);
	}

private:
	inline void putNewline()
	{
		if(nTagMode == TAG_WAP && nTextPos)
			m_Stream.Write(L"<br />\r\n", 16, NULL);
		else if(nLinePos)
			m_Stream.Write(L"\r\n", 4, NULL);

		nLinePos = 0;
		nTextPos = 0;

		for(int i = 0; i < nIndent; i ++)
		{
			m_Stream.Write(L"  ", 4, NULL);
			nLinePos = 2;
		}
	}

	inline void putSpace()
	{
		if(nLinePos > 77 && nTagMode == TAG_HTML)
			putNewline();
		else
		{
			m_Stream.Write(L" ", 2, NULL);
			nLinePos ++;
		}
	}

	inline void putSplit()
	{
		if(bNewLine)
		{
			if(m_Stream.GetLength())putNewline();

			bSpace = 0;
			bNewLine = 0;
		}else if(bSpace)
		{
			if(!bSkipSpace)putSpace();
			bSpace = 0;
		}
		bSkipSpace = 0;
	}

	inline void CloseTag(int i)
	{
		_tagName* pTag = &s_tags[arrayTagStack[i]];
		int j = (int)arrayTagStack.GetCount() - 1;
		int nTag = arrayTagStack[i];

		while(j >= 0 && arrayTagStack[j] != nTag)
		{
			_tagName* pTag1 = &s_tags[arrayTagStack[j]];
			if(pTag1->modes[nTag + 1] == '2' ||
				(pTag1->modes[nTag + 1] != '1' && pTag1->modes[0] == '1'))
				CloseTag(j);
			j --;
		}

		if(!nPRE && pTag->inside)
		{
			bNewLine = 1;
			if(pTag->after && pTag->before)
			{
				nIndent --;
				if(nIndent < 0)nIndent = 0;
			}
		}
		if(!nPRE && !pTag->inside && pTag->before)bSkipSpace = 1;
		if(nTagMode == TAG_HTML || (nTagMode == TAG_WAP && pTag->isWAP))
		{
			if(nTagMode == TAG_WAP && pTag->id == TAG_P)
				nTextPos = 0;

			putSplit();

			m_Stream.Write(L"</", 4, NULL);
			m_Stream.Write(pTag->name, pTag->length * 2, NULL);
			m_Stream.Write(L">", 2, NULL);
			nLinePos += pTag->length + 3;
		}
		if(pTag->id == TAG_PRE)nPRE --;
		if(!nPRE && pTag->after)bNewLine = 1;

		arrayTagStack.RemoveAt(i);
	}

	WCHAR getchar()
	{
		if(buf_leng)
		{
			WCHAR ch;

			ch = *buf_text ++;
			buf_leng --;
			if(!ch)ch = ' ';
			return ch;
		}

		return 0;
	}

	void nextchar()
	{
		buf_leng --;
		buf_text ++;
	}

	static BOOL inline isSpace(WCHAR ch)
	{
		return !ch || ch == ' ' ||  ch == '\t' || ch == '\r' || ch == '\n';
	}

	int normal_lex()
	{
		WCHAR ch;

		ch = getchar();
		if(!ch)
			return 0;

		fht_leng ++;

		if(ch == '>')
			return ch;

		if(ch == '<')
		{
			if(buf_leng >= 3 && !wcsncmp(buf_text, L"!--", 3))
			{
				WCHAR* ptr;
				int leng;

				ptr = buf_text + 3;
				leng = buf_leng - 3;

				while(leng)
				{
					while(leng && (*ptr != '-'))
					{
						leng --;
						ptr ++;
					}

					if(!leng)
						break;

					if(!wcsncmp(ptr, L"-->", 3))
					{
						buf_text = ptr + 3;
						buf_leng = leng - 3;

						return -1;
					}else
					{
						leng --;
						ptr ++;
					}
				}
			}

			if(buf_leng && (*buf_text == '!' || *buf_text == '?'))
			{
				while(buf_leng && *buf_text != '>')
					nextchar();

				getchar();

				return -1;
			}

			if(buf_leng && (iswalpha(*buf_text) || *buf_text == '/'))
			{
				do
				{
					nextchar();
					fht_leng ++;
				}while(buf_leng && !isSpace(*buf_text) && *buf_text != '/' && *buf_text != '>');

				if(fht_leng == 7 && !_wcsnicmp(fht_text + 1, L"script", 6))
				{
					while(buf_leng && *buf_text != '>')
						nextchar();

					getchar();

					while(buf_leng)
					{
						while(buf_leng && *buf_text != '<')
							nextchar();

						if(buf_leng && !_wcsnicmp(buf_text, L"</script", 8) && (isSpace(buf_text[8]) || buf_text[8] == '>'))
						{
							while(buf_leng && *buf_text != '>')
								nextchar();

							getchar();

							break;
						}else getchar();
					}

					return -1;
				}

				if(fht_text[1] == '/')
				{
					while(buf_leng && *buf_text != '>')
						nextchar();

					getchar();

					return TAG_CLOSE_ID;
				}

				mode = TAG_MODE;
				return TAG_START_ID;
			}

			return ch;
		}

		if(isSpace(ch))
		{
			while(buf_leng && isSpace(*buf_text))
			{
				nextchar();
				fht_leng ++;
			}

			return SPACETEXT;
		}

		if(nTagMode == TAG_WAP && ch == '&' && buf_leng >= 5 && !wcsncmp(buf_text, L"nbsp;", 5))
		{
			buf_leng -= 5;
			buf_text += 5;
			fht_leng += 5;

			return NBSPTEXT;
		}

		while(buf_leng && !isSpace(*buf_text) && *buf_text != '&' && *buf_text != '<' && *buf_text != '>')
		{
			nextchar();
			fht_leng ++;
		}

		return PLANTEXT;
	}

	int tag_lex()
	{
		WCHAR ch;

		while(buf_leng && isSpace(*buf_text))
			nextchar();

		if(*buf_text == '<')
		{
			mode = 0;
			return '>';
		}

		fht_text = buf_text;

		ch = getchar();
		if(!ch)
			return 0;

		fht_leng ++;

		switch(ch)
		{
		case '>':mode = 0;
		case '=':return ch;
		};

		if(ch == '\'' || ch == '\"')
		{
			WCHAR* ptr;
			int leng;

			ptr = buf_text;
			leng = buf_leng;

			while(leng && *ptr != '\n' && *ptr != '\r' && *ptr != ch)
			{
				leng --;
				ptr ++;
			}

			fht_leng = buf_leng - leng;
			fht_text ++;
			if(*ptr == ch)
			{
				buf_text = ptr + 1;
				buf_leng = leng - 1;
			}
			return TAG_STR;
		}

		if(bTagValue)
			while(buf_leng && !isSpace(*buf_text) && *buf_text != '>' && *buf_text != '<')
			{
				nextchar();
				fht_leng ++;
			}
		else
		{
			while(buf_leng && !isSpace(*buf_text) && *buf_text != '>' && *buf_text != '<' && *buf_text != '=')
			{
				nextchar();
				fht_leng ++;
			}

			if(fht_leng == 5 && !_wcsnicmp(fht_text, L"style", 5))
			{
				while(buf_leng && isSpace(*buf_text))
					nextchar();

				if(*buf_text != '=')return -1;

				nextchar();

				while(buf_leng && isSpace(*buf_text))
					nextchar();

				if(*buf_text != '\"' && *buf_text != '\'')
				{
					while(buf_leng && isSpace(*buf_text))
						nextchar();
					styleEndChar = ' ';
				}else
				{
					styleEndChar = *buf_text;
					nextchar();
				}

				mode = STYLE_MODE;
				return STYLE_BEGIN;
			}
		}

		return TAG_ID;
	}

	BOOL inline isStyleEnd(WCHAR ch)
	{
		if(styleEndChar == ' ')
			return ch == '>' || isSpace(ch);

		return (ch == styleEndChar);
	}

	int style_lex()
	{
		WCHAR ch;

		if(isStyleEnd(*buf_text))
		{
			if(*buf_text != '>')nextchar();
			mode = TAG_MODE;
			return STYLE_END;
		}

		while(buf_leng && isSpace(*buf_text))
			nextchar();

		fht_text = buf_text;

		ch = getchar();
		if(!ch)
			return 0;

		fht_leng ++;

		if(ch == styleEndChar)
		{
			mode = TAG_MODE;
			return STYLE_END;
		}

		if(ch == ':' || ch == ';')return ch;

		if(bTagValue)
		{
			while(buf_leng && *buf_text != ';' && !isStyleEnd(*buf_text))
			{
				nextchar();
				fht_leng ++;
			}

			while(isSpace(fht_text[fht_leng - 1]))
				fht_leng --;

			return STYLE_VALUE;
		}else
		{
			while(buf_leng && !isSpace(*buf_text) && *buf_text != ':' && *buf_text != ';' && !isStyleEnd(*buf_text))
			{
				nextchar();
				fht_leng ++;
			}

			return STYLE_ID;
		}

		return 0;
	}

	int fht_lex()
	{
		if(!buf_leng)
			return 0;

		fht_text = buf_text;
		fht_leng = 0;

		switch(mode)
		{
		case 0: return normal_lex();
		case TAG_MODE: return tag_lex();
		case STYLE_MODE: return style_lex();
		};

		return 0;
	}

private:
	WCHAR* buf_text;
	int buf_leng;
	WCHAR* fht_text;
	int fht_leng;
	CAtlArray<int> arrayTagStack;
	int bSpace, bNewLine, bSkipSpace, nPRE;
	int nLinePos;
	int nTextPos;
	int nTagMode;
	int nIndent;
	int mode;
	BOOL bTagValue;
	WCHAR styleEndChar;
	CXMemStream m_Stream;
};

STDMETHODIMP CXEncoding::HtmlFormat(BSTR TextString, VARIANT varFormat, BSTR *retVal)
{
	CXString strTemp;
	CXHtmlLex lex(TextString, varGetNumbar(varFormat));

	lex.fmtHTML();
	lex.GetString(strTemp);
	*retVal = strTemp.AllocSysString();

	return S_OK;
}

