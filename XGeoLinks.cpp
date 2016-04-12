// XGeoLinks.cpp : Implementation of CXGeoLinks

#include "stdafx.h"
#include "XGeoLinks.h"

#include <math.h>

// CXGeoLinks

#define GEO_SORT	1
#define GEO_LIMIT	2
#define GEO_MAINTAG	4
#define GEO_EQUAL	8
#define GEO_NOSUB	16


STDMETHODIMP CXGeoLinks::get_halfLife(int* pVal)
{
	*pVal = m_nhalfLifeTimes;
	return S_OK;
}

STDMETHODIMP CXGeoLinks::put_halfLife(int newVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	if(newVal > 0)
		m_nhalfLifeTimes = newVal;

	return S_OK;
}

STDMETHODIMP CXGeoLinks::get_halfRange(int* pVal)
{
	*pVal = m_nhalfLifeRange;
	return S_OK;
}

STDMETHODIMP CXGeoLinks::put_halfRange(int newVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	if(newVal > 0)
		m_nhalfLifeRange = newVal;

	return S_OK;
}

STDMETHODIMP CXGeoLinks::get_Count(int* pVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	*pVal = (int)m_mapNodes.GetCount();

	return S_OK;
}

STDMETHODIMP CXGeoLinks::Modify(IXGeoLinks** pVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	CXComPtr<CXGeoLinks> pModify;

	pModify.CoCreateObject();
	pModify->m_pUpdate = this;

	pModify->m_nhalfLifeTimes = m_nhalfLifeTimes;
	pModify->m_nhalfLifeRange = m_nhalfLifeRange;
	pModify->m_baseTime = m_baseTime;

	return pModify.QueryInterface(pVal);
}

//-----------------------------------------------------------------------------------


CXGeoLinks::CXTree::~CXTree()
{
	int i;

	for(i = 0; i < 26; i ++)
		delete m_aSubTrees[i];

	POSITION pos;
	CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair* pPair;

	pos = m_aTagNodes.GetHeadPosition();
	while(pos)
	{
		pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.GetNext(pos);
		delete pPair->m_value;
	}

	pos = m_aTagNodes1.GetHeadPosition();
	while(pos)
	{
		pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes1.GetNext(pos);
		delete pPair->m_value;
	}
}

void CXGeoLinks::CXTree::AddAreaItem(CXNode* pNode, LPCWSTR mark)
{
	int n, i;

	if(mark[0])
	{
		n = (mark[0] & 0x1f) - 1;

		CXTree* pTree = m_aSubTrees[n];

		if(!pTree)
			pTree = m_aSubTrees[n] = new CXTree();

		pTree->AddAreaItem(pNode, mark + 1);
	}else
	{
		CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair* pPair;
		for(i = 0; i < (int)pNode->m_strTags.GetCount(); i ++)
		{
			if(i == 0)
			{
				pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.Lookup(pNode->m_strTags[0]);

				if(pPair == NULL)
					pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.SetAt(pNode->m_strTags[0], new CAtlList< CXComPtr<CXNode> >);
			}else
			{
				pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes1.Lookup(pNode->m_strTags[i]);

				if(pPair == NULL)
					pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes1.SetAt(pNode->m_strTags[i], new CAtlList< CXComPtr<CXNode> >);
			}

			POSITION pos, pos1;
			CXNode* pNode1;

			pos = pos1 = pPair->m_value->GetHeadPosition();
			while(pos)
			{
				pos1 = pos;
				pNode1 = pPair->m_value->GetNext(pos);

				if(pNode1->m_nLife < pNode->m_nLife)
				{
					pPair->m_value->InsertBefore(pos1, pNode);
					break;
				}
				pos1 = NULL;
			}

			if(!pos1)
				pPair->m_value->AddTail(pNode);
		}
	}
}

static int _cmpcityid(const void* p1, const void* p2)
{
	return *(const int*)p1 - *(const int*)p2;
}

void CXGeoLinks::CXRoot::AddAreaItem(CXNode* pNode)
{
	CityInfoStruct* pCity = (CityInfoStruct*)bsearch(&pNode->m_nCityID, s_cityname, CITYCOUNT, sizeof(s_cityname[0]), _cmpcityid);

	if(pCity == NULL)
		return;
	int n = pCity - s_cityname;

	CXTree* pTree = m_mapCitys[n];

	if(!pTree)
		pTree = m_mapCitys[n] = new CXTree();

	pTree->AddAreaItem(pNode, pNode->m_strMark);
}

//-----------------------------------------------------------------------------------

STDMETHODIMP CXGeoLinks::Commit(void)
{
	if(!m_pUpdate)
		return DISP_E_BADCALLEE;

	POSITION pos;
	CRBMap< int, CXComPtr<CXNode> >::CPair* pPair;

	m_pUpdate->m_cs.Enter();
	pos = m_pUpdate->m_mapNodes.GetHeadPosition();
	while(pos)
	{
		pPair = (CRBMap< int, CXComPtr<CXNode> >::CPair*)m_pUpdate->m_mapNodes.GetNext(pos);
		if(pPair->m_value->m_nLife > 0)
			m_root->AddAreaItem(pPair->m_value);
		else m_pUpdate->m_mapNodes.RemoveAt(pPair);
	}
	m_pUpdate->m_cs.Leave();

	m_pUpdate->m_root = m_root;

	return S_OK;
}

void CXGeoLinks::AddOneItem(int id, int cityid, double Latitude, double Longitude, LPCWSTR mark, LPCWSTR tag, double life, DATE time, LPCWSTR value)
{
	int i, n;
	CXComPtr<CXNode> pNode;

	while(Latitude < -90)Latitude += 180;
	while(Latitude > 90)Latitude -= 180;
	while(Longitude < -180)Longitude += 360;
	while(Longitude > 180)Longitude -= 360;

	pNode.CreateObject();
	pNode->m_nID = id;
	pNode->m_nCityID = cityid;

	for(i = 0; tag[i];)
	{
		for(; tag[i] == ' '; i ++);

		n = i;
		for(; tag[i] && tag[i] != ' '; i ++);
		if(i > n)
			pNode->m_strTags.Add(CXString(tag + n, i - n));
	}

	if(pNode->m_strTags.GetCount() == 0)
		pNode->m_strTags.Add(L"");

	pNode->m_nlat = Latitude;
	pNode->m_nlng = Longitude;
	pNode->m_strValue = value;

	n = wcslen(mark);
	if(n > 15)return;
	wcsncpy_s(pNode->m_strMark, 16, mark, n);

	for(i = 0; i < n; i ++)
		if(pNode->m_strMark[i] >= 'A' && pNode->m_strMark[i] <= 'Z')
			pNode->m_strMark[i] += 'a' - 'A';
		else if(pNode->m_strMark[i] < 'a' || pNode->m_strMark[i] > 'z')
			return;

	pNode->m_nLife = life * pow(2, (time - m_baseTime) * 24 / m_nhalfLifeTimes);
	m_pUpdate->m_mapNodes.SetAt(id, pNode);
}

STDMETHODIMP CXGeoLinks::AddItem(int id, int cityid, double Latitude, double Longitude, int type, BSTR mark, BSTR tag, double life, DATE time, BSTR value)
{
	if(m_pUpdate == NULL)
		return DISP_E_BADCALLEE;

	if(type < 0 || type > 25 || id == 0 || cityid == 0 || life <= 0)
		return DISP_E_BADCALLEE;

	m_pUpdate->m_cs.Enter();
	AddOneItem(id, cityid, Latitude, Longitude, CXString(WCHAR('a' + type)) + mark, tag, life, time, value);
	m_pUpdate->m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXGeoLinks::Remove(int id)
{
	CRBMap< int, CXComPtr<CXNode> >::CPair* pPair;
	CXGeoLinks* pLinks = m_pUpdate;
	
	if(pLinks == NULL)
		pLinks = this;

	pLinks->m_cs.Enter();
	if(pPair = pLinks->m_mapNodes.Lookup(id))
	{
		pPair->m_value->m_nLife = 0;
		pLinks->m_mapNodes.RemoveAt(pPair);
	}
	pLinks->m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXGeoLinks::RemoveAll()
{
	if(m_pUpdate == NULL)
		return DISP_E_BADCALLEE;

	m_pUpdate->m_mapNodes.RemoveAll();
	m_pUpdate->m_baseTime = m_baseTime = GetVariantTime();

	return S_OK;
}

void inline getSplitString(BSTR& batchText, CXString& str)
{
	BSTR strTemp = batchText;
	WCHAR ch = 0;

	while((ch = *batchText) && (ch != '\t') && (ch != '\n') && (ch != '\r'))
		batchText ++;

	str.SetString(strTemp, batchText - strTemp);

	if(ch == '\t')batchText ++;
}

void inline getLeftString(BSTR& batchText, CXString& str)
{
	BSTR strTemp = batchText;
	WCHAR ch = 0;

	while((ch = *batchText) && (ch != '\n') && (ch != '\r'))
		batchText ++;

	str.SetString(strTemp, batchText - strTemp);
}

STDMETHODIMP CXGeoLinks::BatchAdd(BSTR batchText)
{
	if(m_pUpdate == NULL)
		return DISP_E_BADCALLEE;

	HRESULT hr = S_OK;
	CXString mark, tag, value, str;
	double Latitude, Longitude;
	DATE time;
	int id, cityid, type;
	double life;

	CXAutoLock l(m_pUpdate->m_cs, TRUE);

	while(*batchText)
	{
		while(*batchText == '\n' || (*batchText == '\r'))
			batchText ++;

		getSplitString(batchText, str);
		id = _wtol(str);

		getSplitString(batchText, str);
		cityid = _wtol(str);

		getSplitString(batchText, str);
		Latitude = _wtof(str);
		getSplitString(batchText, str);
		Longitude = _wtof(str);

		getSplitString(batchText, str);
		type = _wtol(str);
		getSplitString(batchText, mark);
		getSplitString(batchText, tag);

		getSplitString(batchText, str);
		life = _wtof(str);

		getSplitString(batchText, str);
		hr = ::VarDateFromStr((OLECHAR*)(LPCWSTR)str, 0, 0, &time);
		if(FAILED(hr))break;

		getLeftString(batchText, value);

		if(id && cityid && !value.IsEmpty() && !mark.IsEmpty())
		{
			if(type < 0 || type > 25)
				return DISP_E_BADCALLEE;

			if(life > 0)
				AddOneItem(id, cityid, Latitude, Longitude, WCHAR('a' + type) + mark, tag, life, time, value);
			else m_pUpdate->m_mapNodes.RemoveKey(id);
		}
	}

	return S_OK;
}

static double s_weaken[] = {985.1853368,942.8660343,878.7397112,800.4109404,715.5417528,630.5095042,549.8200809,476.139518,410.6597493,353.5533906,304.3768299,262.3706556,226.6582754,196.3642546,170.6769835,148.8761067,130.3393642,114.5384272,101.029595,89.4427191,79.47016273,70.85666854,63.39047876,56.89576695,51.22630019,46.26019063,41.89557971,38.04710373,34.64300567,31.6227766,28.93522902,26.5369211,24.39086588,22.46547166,20.73367003,19.17219655,17.76099558,16.48272664,15.3223535,14.26680147,13.30467042,12.42599395,11.62203667,10.88512297,10.20849213,9.586175286,9.012890622,8.483953998,7.995202399,7.542928275,7.123823055,6.734928447,6.373594335,6.037442314,5.724334022,5.432343587,5.159733592,4.904934079,4.666524158,4.443215873,4.233840017,4.037333642,3.852729029,3.679143953,3.515773046,3.361880149,3.216791507,3.079889723,2.950608359,2.828427125,2.712867571,2.603489237,2.499886191,2.401683931,2.308536595,2.220124448,2.136151617,2.056344049,1.980447665,1.908226686,1.839462123,1.7739504,1.711502112,1.651940886,1.595102356,1.54083322,1.488990385,1.439440184,1.392057663,1.346725928,1.303335543,1.261783988,1.221975153,1.183818874,1.147230515,1.112130574,1.078444324,1.046101484,1.015035915,0.985185337,0.956491069,0.928897792,0.902353323,0.876808416,0.852216563,0.828533827,0.805718672,0.783731814,0.76253608,0.742096279,0.722379075,0.703352879,0.68498774,0.667255251,0.650128453,0.633581751,0.617590835,0.602132606,0.587185105,0.572727447,0.558739765,0.545203146,0.532099585,0.519411928,0.50712383,0.495219711,0.483684711,0.472504655,0.461666015,0.451155876,0.440961904,0.431072316,0.421475852,0.412161746,0.403119703,0.394339874,0.385812837,0.377529568,0.369481432,0.361660152,0.354057803,0.346666787,0.339479819,0.332489915,0.325690375,0.319074768,0.312636926,0.306370923,0.300271071,0.294331905,0.288548174,0.282914831,0.277427026,0.272080093,0.266869546,0.261791068,0.256840506,0.252013861,0.247307283,0.242717067,0.238239641,0.233871566,0.229609526,0.225450328,0.221390889,0.21742824,0.213559515,0.209781952,0.206092881,0.202489731,0.198970015,0.195531335,0.192171374,0.188887894,0.185678732,0.182541801,0.17947508,0.176476618,0.173544527,0.170676983,0.167872221,0.165128533,0.162444267,0.159817824,0.157247656,0.154732266,0.152270203,0.149860062,0.147500483,0.145190147,0.142927777,0.140712137,0.138542025,0.136416281,0.134333776,0.132293417,0.130294144,0.128334929,0.126414774,0.124532711,0.1226878,0.120879131,0.119105818,0.117367002,0.115661849,0.113989547,0.112349312,0.110740377,0.109162,0.10761346,0.106094056,0.104603105,0.103139945,0.101703932,0.10029444,0.098910859,0.097552598,0.096219079,0.094909743,0.093624043,0.09236145,0.091121446,0.08990353,0.088707211,0.087532013,0.086377473,0.085243138,0.08412857,0.083033338,0.081957027,0.080899229,0.079859548,0.078837599,0.077833004,0.076845398,0.075874423,0.07491973,0.07398098,0.073057841,0.072149991,0.071257114,0.070378903,0.069515058,0.068665287,0.067829305,0.067006832,0.066197599,0.065401338,0.064617791,0.063846707,0.063087837,0.062340941,0.061605785,0.060882138,0.060169776,0.05946848,0.058778037,0.058098236,0.057428874,0.056769752,0.056120674,0.05548145,0.054851893,0.054231822,0.053621059,0.05301943,0.052426764,0.051842896,0.051267662,0.050700905,0.050142468,0.049592199,0.04904995,0.048515575,0.047988931,0.04746988,0.046958285,0.046454013,0.045956933,0.045466918,0.044983842,0.044507583,0.044038022,0.043575041,0.043118525,0.042668361,0.042224441,0.041786655,0.041354899,0.040929069,0.040509063,0.040094783,0.039686131,0.039283012,0.038885332,0.038493001,0.038105929,0.037724027,0.03734721,0.036975394,0.036608496,0.036246435,0.035889131,0.035536506,0.035188485,0.034844992,0.034505955,0.0341713,0.033840958,0.033514859,0.033192935,0.03287512,0.032561348,0.032251556,0.03194568,0.031643659,0.031345432,0.03105094,0.030760125,0.030472929,0.030189297,0.029909172,0.029632502,0.029359233,0.029089313,0.02882269,0.028559316,0.02829914,0.028042113,0.02778819,0.027537322,0.027289464,0.027044571,0.026802599,0.026563504,0.026327244,0.026093777,0.025863061,0.025635057,0.025409725,0.025187025,0.024966919,0.024749371,0.024534341,0.024321795,0.024111697,0.023904011,0.023698703,0.023495738,0.023295085,0.023096709,0.022900579,0.022706662,0.022514929,0.022325348,0.022137888,0.021952522,0.021769218,0.021587949,0.021408687,0.021231404,0.021056072,0.020882665,0.020711157,0.020541522,0.020373733,0.020207767,0.020043598,0.019881202,0.019720556,0.019561635,0.019404417,0.019248879,0.019094999,0.018942754,0.018792123,0.018643084,0.018495618,0.018349702,0.018205316,0.018062442,0.017921058,0.017781145,0.017642685,0.017505659,0.017370048,0.017235833,0.017102998,0.016971524,0.016841393,0.01671259,0.016585096,0.016458896,0.016333973,0.01621031,0.016087893,0.015966704,0.01584673,0.015727955,0.015610363,0.015493941,0.015378673,0.015264545,0.015151544,0.015039656,0.014928866,0.014819161,0.014710529,0.014602956,0.014496428,0.014390935,0.014286462,0.014182998,0.01408053,0.013979047,0.013878537,0.013778988,0.013680388,0.013582727,0.013485993,0.013390176,0.013295263,0.013201246,0.013108113,0.013015853,0.012924458,0.012833916,0.012744217,0.012655353,0.012567312,0.012480087,0.012393666,0.012308041,0.012223204,0.012139144,0.012055853,0.011973322,0.011891542,0.011810506,0.011730204,0.011650628,0.011571771,0.011493623,0.011416178,0.011339426,0.011263361,0.011187975,0.01111326,0.011039208,0.010965813,0.010893067,0.010820963,0.010749494,0.010678653,0.010608433,0.010538827,0.010469828,0.010401431,0.010333628,0.010266413,0.010199779,0.010133721,0.010068232,0.010003306,0.009938937,0.009875119,0.009811846,0.009749112,0.009686912,0.00962524,0.009564091,0.009503458,0.009443337,0.009383721,0.009324606,0.009265987,0.009207858,0.009150214,0.00909305,0.009036362,0.008980143,0.00892439,0.008869097,0.008814261,0.008759875,0.008705935,0.008652438,0.008599378,0.008546751,0.008494552,0.008442778,0.008391423,0.008340484,0.008289957,0.008239836,0.008190119,0.008140801,0.008091878,0.008043347,0.007995202,0.007947441,0.00790006,0.007853055,0.007806421,0.007760156,0.007714256,0.007668718,0.007623536,0.007578709,0.007534233,0.007490104,0.007446319,0.007402875,0.007359768,0.007316995,0.007274552,0.007232438,0.007190647,0.007149178,0.007108027,0.007067192,0.007026668,0.006986454,0.006946546,0.006906942,0.006867637,0.006828631,0.006789919,0.006751499,0.006713369,0.006675525,0.006637965,0.006600686,0.006563686,0.006526962,0.006490511,0.006454332,0.00641842,0.006382775,0.006347393,0.006312271,0.006277409,0.006242803,0.00620845,0.006174349,0.006140498,0.006106893,0.006073534,0.006040416,0.006007539,0.0059749,0.005942497,0.005910328,0.005878391,0.005846684,0.005815204,0.005783949,0.005752918,0.005722109,0.005691519,0.005661147,0.005630991,0.005601048,0.005571318,0.005541797,0.005512485,0.005483378,0.005454477,0.005425778,0.00539728,0.005368981,0.00534088,0.005312975,0.005285263,0.005257744,0.005230416,0.005203277,0.005176325,0.005149559,0.005122977,0.005096578,0.005070359,0.005044321,0.00501846,0.004992776,0.004967267,0.004941931,0.004916767,0.004891774,0.00486695,0.004842293,0.004817803,0.004793477,0.004769316,0.004745316,0.004721477,0.004697797,0.004674275,0.004650911,0.004627701,0.004604646,0.004581744,0.004558993,0.004536393,0.004513941,0.004491638,0.004469482,0.00444747,0.004425603,0.00440388,0.004382298,0.004360857,0.004339555,0.004318392,0.004297367,0.004276478,0.004255724,0.004235104,0.004214616,0.004194261,0.004174037,0.004153943,0.004133977,0.004114139,0.004094428,0.004074842,0.004055381,0.004036044,0.00401683,0.003997737,0.003978765,0.003959913,0.00394118,0.003922565,0.003904067,0.003885685,0.003867419,0.003849266,0.003831227,0.003813301,0.003795486,0.003777782,0.003760188,0.003742703,0.003725326,0.003708057,0.003690894,0.003673837,0.003656885,0.003640037,0.003623292,0.00360665,0.00359011,0.00357367,0.003557331,0.003541091,0.00352495,0.003508907,0.003492961,0.003477112,0.003461358,0.0034457,0.003430135,0.003414664,0.003399287,0.003384001,0.003368807,0.003353703,0.00333869,0.003323766,0.003308931,0.003294185,0.003279525,0.003264953,0.003250467,0.003236066,0.00322175,0.003207519,0.003193371,0.003179306,0.003165324,0.003151424,0.003137605,0.003123867,0.003110208,0.003096629,0.00308313,0.003069708,0.003056364,0.003043098,0.003029908,0.003016794,0.003003756,0.002990793,0.002977904,0.00296509,0.002952348,0.00293968,0.002927084,0.00291456,0.002902107,0.002889725,0.002877413,0.002865171,0.002852999,0.002840895,0.00282886,0.002816893,0.002804993,0.00279316,0.002781393,0.002769692,0.002758057,0.002746487,0.002734982,0.002723541,0.002712163,0.002700849,0.002689598,0.002678409,0.002667282,0.002656216,0.002645212,0.002634268,0.002623385,0.002612562,0.002601798,0.002591093,0.002580446,0.002569858,0.002559328,0.002548855,0.00253844,0.002528081,0.002517778,0.002507531,0.00249734,0.002487203,0.002477122,0.002467095,0.002457122,0.002447203,0.002437337,0.002427524,0.002417764,0.002408055,0.002398399,0.002388794,0.002379241,0.002369738,0.002360286,0.002350884,0.002341533,0.00233223,0.002322977,0.002313772,0.002304617,0.002295509,0.002286449,0.002277437,0.002268473,0.002259555,0.002250684,0.002241859,0.002233081,0.002224348,0.00221566,0.002207018,0.002198421,0.002189868,0.00218136,0.002172895,0.002164475,0.002156097,0.002147763,0.002139472,0.002131224,0.002123018,0.002114853,0.002106731,0.00209865,0.002090611,0.002082612,0.002074655,0.002066738,0.002058861,0.002051023,0.002043226,0.002035468,0.00202775,0.00202007,0.002012429,0.002004827,0.001997263,0.001989736,0.001982248,0.001974797,0.001967384,0.001960007,0.001952667,0.001945364,0.001938098,0.001930867,0.001923672,0.001916513,0.00190939,0.001902302,0.001895248,0.00188823,0.001881246,0.001874297,0.001867382,0.001860501,0.001853653,0.001846839,0.001840059,0.001833311,0.001826597,0.001819915,0.001813266,0.001806649,0.001800065,0.001793512,0.001786991,0.001780502,0.001774044,0.001767617,0.001761221,0.001754857,0.001748522,0.001742218,0.001735945,0.001729701,0.001723488,0.001717304,0.00171115,0.001705025,0.001698929,0.001692862,0.001686824,0.001680815,0.001674834,0.001668882,0.001662958,0.001657062,0.001651193,0.001645353,0.001639539,0.001633754,0.001627995,0.001622263,0.001616559,0.001610881,0.001605229,0.001599604,0.001594005,0.001588433,0.001582886,0.001577365,0.00157187,0.0015664,0.001560956,0.001555536,0.001550142,0.001544773,0.001539428,0.001534108,0.001528813,0.001523542,0.001518295,0.001513072,0.001507873,0.001502698,0.001497547,0.001492419,0.001487314,0.001482233,0.001477175,0.00147214,0.001467128,0.001462138,0.001457171,0.001452226,0.001447304,0.001442404,0.001437526,0.001432671,0.001427837,0.001423024,0.001418234,0.001413464,0.001408716,0.00140399,0.001399284,0.0013946,0.001389936,0.001385293,0.001380671,0.001376069,0.001371488,0.001366927,0.001362386,0.001357866,0.001353365,0.001348884,0.001344423,0.001339982,0.00133556,0.001331158,0.001326775,0.001322411,0.001318066,0.00131374,0.001309433,0.001305145,0.001300876,0.001296625,0.001292393,0.001288179,0.001283984,0.001279807,0.001275647,0.001271506,0.001267383,0.001263277,0.001259189,0.001255119,0.001251067,0.001247031,0.001243013,0.001239013,0.001235029,0.001231063,0.001227113,0.001223181,0.001219265,0.001215366,0.001211483,0.001207617,0.001203768,0.001199935,0.001196118,0.001192317,0.001188532,0.001184763,0.001181011,0.001177274,0.001173552,0.001169847,0.001166157,0.001162482,0.001158823,0.00115518,0.001151551,0.001147938,0.00114434,0.001140757,0.001137189,0.001133635,0.001130097,0.001126573,0.001123064,0.001119569,0.001116089,0.001112623,0.001109172,0.001105735,0.001102312,0.001098903,0.001095508,0.001092127,0.00108876,0.001085407,0.001082068,0.001078742,0.00107543,0.001072132,0.001068846,0.001065575,0.001062317,0.001059072,0.00105584,0.001052621,0.001049415,0.001046223,0.001043043,0.001039876,0.001036722,0.001033581,0.001030452,0.001027336,0.001024233,0.001021142,0.001018064,0.001014997,0.001011944,0.001008902,0.001005873,0.001002855,0.000999855};

void CXGeoLinks::CXTree::GetTreeNodes(LPCWSTR tag, CAtlArray<CXString>* keys, int rate, int mode, CAtlArray<CXNodeIndex>& nodes, int nCount)
{
	int i;
	CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair* pPair;
	CXTree* pTree;

	if(*tag)
	{
		if(mode & GEO_MAINTAG)
		{
			if(pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes1.Lookup(tag))
				FillNodes(pPair->m_value, keys, mode, rate, nodes, nCount);
		}else
		{
			if(pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.Lookup(tag))
				FillNodes(pPair->m_value, keys, mode, rate, nodes, nCount);
		}
	}else
	{
		POSITION pos;

		pos = m_aTagNodes.GetHeadPosition();
		while(pos)
		{
			pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.GetNext(pos);
			FillNodes(pPair->m_value, keys, mode, rate, nodes, nCount);
		}
	}

	for(i = 0; i < 26; i ++)
		if(pTree = m_aSubTrees[i])
			pTree->GetTreeNodes(tag, keys, rate, mode, nodes, nCount);
}

void CXGeoLinks::CXTree::GetNodes(LPCWSTR mark, LPCWSTR tag, CAtlArray<CXString>* keys, int mode, int *rates, CAtlArray<CXNodeIndex>& nodes, int nCount)
{
	CXTree* pTree;
	LPCWSTR mark1;

	int i;

	for(i = 0; i < 26; i ++)
		if(rates[i] && (pTree = m_aSubTrees[i]))
		{
			mark1 = mark;
			WCHAR ch = *mark1;

			while(ch)
			{
				mark1 ++;
				ch = (ch & 0x1f) - 1;

				if(!pTree->m_aSubTrees[ch])
				{
					pTree = NULL;
					break;
				}

				pTree = pTree->m_aSubTrees[ch];
				ch = *mark1;
			}

			if(pTree)
				pTree->GetTreeNodes(tag, keys, rates[i], mode, nodes, nCount);
		}
}

void CXGeoLinks::CXTree::FillNodes(CAtlList< CXComPtr<CXNode> >* pList, CAtlArray<CXString>* keys, int mode, int rate, CAtlArray<CXNodeIndex>& nodes, int nCount)
{
	CXNodeIndex node;
	POSITION pos;
	int i, n;
	int KeyCount = 0;
	LPCWSTR p, p1;
	double r = 1;

	if(keys)
		KeyCount = keys->GetCount();

	pos = pList->GetHeadPosition();
	n = 0;
	while(pos && n <= nCount)
	{
		node.m_pNode = pList->GetNext(pos);

		if(KeyCount)
		{
			r = 0;

			p = node.m_pNode->m_strValue;

			for(i = 0; i < KeyCount; i ++)
				if(p1 = StrStrIW(p, keys->GetAt(i)))
					r += 16.0 / ((p1 - p) / 4 + 1);

			if(r == 0)
				continue;
		}

		if(node.m_pNode->m_nLife == 0)
			continue;

		if((mode & GEO_SORT) == 0)
			node.m_nLife = pow(node.m_pNode->m_nLife * 100, 0.1) * rand() * rate * r;
		else
			node.m_nLife = node.m_pNode->m_nLife * r;

		for(i = n; (i < (int)nodes.GetCount()) && (node.m_nLife < nodes[i].m_nLife); i ++);
		if(i < nCount)
		{
			nodes.InsertAt(i, node);
			if(!KeyCount && (mode & GEO_SORT))n = i + 1;
		}

		if((int)nodes.GetCount() > nCount)
			nodes.RemoveAt(nCount, nodes.GetCount() - nCount);
	}
}

void CXGeoLinks::CXRoot::GetNodes(int cityid, double Latitude, double Longitude, LPCWSTR mark, LPCWSTR tag, CAtlArray<CXString>* keys, int mode, int nCount, int *rates, CAtlArray<CXNodeIndex>& nodes)
{
	CAtlArray<CXNodeIndex> nodes1;
	CAtlArray<CXNodeIndex> nodes2;
	int nCount1 = nCount;
	int p1, p2;
	int cityid1;

	CityInfoStruct* pCity = (CityInfoStruct*)bsearch(&cityid, s_cityname, CITYCOUNT, sizeof(s_cityname[0]), _cmpcityid);

	if(pCity == NULL)
		return;

	p1 = pCity - s_cityname;
	p2 = p1 + 1;
	cityid1 = mode & GEO_NOSUB ? cityid : nextcity(cityid);

	while(nCount1 > 0 && ((p1 >= 0 && s_cityname[p1].id >= cityid) || (p2 < CITYCOUNT && s_cityname[p2].id < cityid1)))
	{
		while(p1 >= 0 && s_cityname[p1].id >= cityid)
		{
			if(m_mapCitys[p1] && (!(mode & GEO_NOSUB) || s_cityname[p1].id == cityid))
			{
				if(mode & GEO_EQUAL)
				{
					m_mapCitys[p1]->GetNodes(mark, tag, keys, mode & ~GEO_MAINTAG, rates, nodes, nCount);
					if(*tag && (mode & GEO_MAINTAG))
						m_mapCitys[p1]->GetNodes(mark, tag, keys, mode, rates, nodes, nCount);
				}else
				{
					m_mapCitys[p1]->GetNodes(mark, tag, keys, mode & ~GEO_MAINTAG, rates, nodes1, nCount1);
					if(*tag && (mode & GEO_MAINTAG))
						m_mapCitys[p1]->GetNodes(mark, tag, keys, mode, rates, nodes2, nCount1);
				}
			}

			p1 --;
		}

		if(!(mode & GEO_NOSUB))
		{
			while(p2 < CITYCOUNT && s_cityname[p2].id < cityid1)
			{
				if(p2 < CITYCOUNT && m_mapCitys[p2])
				{
					if(mode & GEO_EQUAL)
					{
						m_mapCitys[p2]->GetNodes(mark, tag, keys, mode & ~GEO_MAINTAG, rates, nodes, nCount);
						if(*tag && (mode & GEO_MAINTAG))
							m_mapCitys[p2]->GetNodes(mark, tag, keys, mode, rates, nodes, nCount);
					}else
					{
						m_mapCitys[p2]->GetNodes(mark, tag, keys, mode & ~GEO_MAINTAG, rates, nodes1, nCount1);
						if(*tag && (mode & GEO_MAINTAG))
							m_mapCitys[p2]->GetNodes(mark, tag, keys, mode, rates, nodes2, nCount1);
					}
				}

				p2 ++;
			}
		}

		if(!(mode & GEO_EQUAL))
		{
			nodes.Append(nodes1);
			nCount1 = nCount - nodes.GetCount();

			if(nCount1 < (int)nodes2.GetCount())
				nodes2.SetCount(nCount1);
			nodes.Append(nodes2);

			nodes1.RemoveAll();
			nodes2.RemoveAll();

			nCount1 = nCount - nodes.GetCount();
		}

		if(mode & GEO_LIMIT)break;

		cityid = parentcity(cityid);
		if(!(mode & GEO_NOSUB))cityid1 = nextcity(cityid);
	}
}

STDMETHODIMP CXGeoLinks::GetList(int cityid, double Latitude, double Longitude, BSTR mark, BSTR tag, BSTR key, int nStart, int nCount, BSTR Types, int mode, VARIANT seed, VARIANT* pVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	CAtlArray<CXNodeIndex> nodes;
	int rates[26];
	int i, l, n;
	CXComPtr<CXRoot> root;
	CAtlArray<CXString> keys;
	LPCWSTR p1, p2;

	if(n = varGetNumbar(seed))
		srand(n);

	n = 0;
	l = ::SysStringLen(Types);
	for(i = 0; i < 26; i ++)
	{
		if(i >= l)
			rates[i] = 0;
		else n += rates[i] = (Types[i] & 0x1f) - 1;
	}

	if(n > 0)
	{
		n = ::SysStringLen(mark);
		if(n > 15)return E_INVALIDARG;

		for(i = 0; i < n; i ++)
			if(!((mark[i] >= 'A' && mark[i] <= 'Z') || (mark[i] >= 'a' && mark[i] <= 'z')))
				return E_INVALIDARG;

		p2 = key;
		while(*p2)
		{
			while(*p2 == ' ')p2 ++;
			p1 = p2;
			while(*p2 && *p2 != ' ')p2 ++;

			if(p2 > p1)
				keys.Add(CXString(p1, p2 - p1));
		}

		root = m_root;
		root->GetNodes(cityid, Latitude, Longitude, mark, tag, &keys, mode, nStart + nCount, rates, nodes);
	}

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;

	if((int)nodes.GetCount() > nStart)
	{
		hr = bstrArray.Create(nodes.GetCount() - nStart);
		if(FAILED(hr))return hr;

		pVar = (VARIANT*)bstrArray.m_psa->pvData;

		for(i = nStart; i < (int)nodes.GetCount(); i ++)
		{
			pVar->vt = VT_BSTR;
			pVar->bstrVal = nodes[i].m_pNode->m_strValue.AllocSysString();
			pVar ++;
		}
	}else
	{
		hr = bstrArray.Create();
		if(FAILED(hr))return hr;
	}

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

void CXGeoLinks::CXRoot::GetCityNodes(int cityid, LPCWSTR mark, LPCWSTR tag, int nCount, int *rates, CAtlArray<CXNodeIndex>& nodes)
{
	CAtlArray<CXNodeIndex> nodes1, nodes2;
	int nCount1 = nCount;
	int p1, p2, i, n;
	int cityid1, cityid2;

	if(cityid != 142499999 && (cityid % 10000))
		cityid = cityid / 10000 * 10000;

	CityInfoStruct* pCity = (CityInfoStruct*)bsearch(&cityid, s_cityname, CITYCOUNT, sizeof(s_cityname[0]), _cmpcityid);

	if(pCity == NULL)
		return;

	p1 = p2 = pCity - s_cityname;
	cityid1 = nextcity(cityid);

	while(nCount1 > 0 && ((p1 >= 0 && s_cityname[p1].id > cityid) || (p2 < CITYCOUNT && s_cityname[p2].id < cityid1)))
	{
		while(p1 >= 0 && s_cityname[p1].id > cityid)
		{
			do
			{
				p1 --;

				if(s_cityname[p1].level < 4)
					break;

				if(p1 >= 0 && m_mapCitys[p1])
				{
					m_mapCitys[p1]->GetNodes(mark, tag, NULL, 1, rates, nodes1, 1);
					if(*tag)
						m_mapCitys[p1]->GetNodes(mark, tag, NULL, 5, rates, nodes1, 1);
				}
			}while(p1 >= 0 && s_cityname[p1].level > 4);

			if(nodes1.GetCount() > 0)
			{
				if(p1 >= 0 && s_cityname[p1].level == 4)
				{
					nodes1[0].m_cityid = s_cityname[p1].id;
					nodes2.Add(nodes1[0]);
				}
				nodes1.RemoveAll();
			}
		}

		while(p2 < CITYCOUNT && s_cityname[p2].id < cityid1)
		{
			do
			{
				if(s_cityname[p2].id == 142321300)
					i = 100;

				if(s_cityname[p2].level < 4)
				{
					p2 ++;
					break;
				}

				if(s_cityname[p2].level == 4)
					cityid2 = s_cityname[p2].id;

				if(m_mapCitys[p2])
				{
					m_mapCitys[p2]->GetNodes(mark, tag, NULL, 1, rates, nodes1, 1);
					if(*tag)
						m_mapCitys[p2]->GetNodes(mark, tag, NULL, 5, rates, nodes1, 1);
				}

				p2 ++;
			}while(p2 < CITYCOUNT && s_cityname[p2].level > 4);

			if(nodes1.GetCount() > 0)
			{
				nodes1[0].m_cityid = cityid2;
				nodes2.Add(nodes1[0]);
				nodes1.RemoveAll();
			}
		}

		nCount1 = nodes.GetCount();
		for(n = 0; n < (int)nodes2.GetCount(); n ++)
		{
			double life = nodes2[n].m_nLife;

			for(i = nCount1; (i < (int)nodes.GetCount()) && (life < nodes[i].m_nLife); i ++);
			if(i < nCount)
			{
				nodes.InsertAt(i, nodes2[n]);

				if((int)nodes.GetCount() > nCount)
					nodes.RemoveAt(nCount, nodes.GetCount() - nCount);
			}
		}
		nodes2.RemoveAll();

		nCount1 = nCount - nodes.GetCount();

		cityid = parentcity(cityid);
		cityid1 = nextcity(cityid);
	}
}

STDMETHODIMP CXGeoLinks::GetCityList(int cityid, BSTR mark, BSTR tag, int nStart, int nCount, BSTR Types, VARIANT* pVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	CAtlArray<CXNodeIndex> nodes;
	int rates[26];
	int i, l, n;
	CXComPtr<CXRoot> root;

	n = 0;
	l = ::SysStringLen(Types);
	for(i = 0; i < 26; i ++)
	{
		if(i >= l)
			rates[i] = 0;
		else n += rates[i] = (Types[i] & 0x1f) - 1;
	}

	if(n > 0)
	{
		n = ::SysStringLen(mark);
		if(n > 15)return E_INVALIDARG;

		for(i = 0; i < n; i ++)
			if(!((mark[i] >= 'A' && mark[i] <= 'Z') || (mark[i] >= 'a' && mark[i] <= 'z')))
				return E_INVALIDARG;

		root = m_root;
		root->GetCityNodes(cityid, mark, tag, nStart + nCount, rates, nodes);
	}

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	CXNode* pNode;

	if((int)nodes.GetCount() > nStart)
	{
		hr = bstrArray.Create(nodes.GetCount() - nStart);
		if(FAILED(hr))return hr;

		pVar = (VARIANT*)bstrArray.m_psa->pvData;

		for(i = nStart; i < (int)nodes.GetCount(); i ++)
		{
			pNode = nodes[i].m_pNode;
			pVar->vt = VT_BSTR;
			pVar->bstrVal = ::SysAllocStringLen(NULL, pNode->m_strValue.GetLength() + 10);
			_itow_s(nodes[i].m_cityid, pVar->bstrVal, 10, 10);
			pVar->bstrVal[9] = '\t';
			CopyMemory(pVar->bstrVal + 10, pNode->m_strValue, pNode->m_strValue.GetLength() * 2);
			pVar ++;
		}
	}else
	{
		hr = bstrArray.Create();
		if(FAILED(hr))return hr;
	}

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

void CXGeoLinks::CXTree::FillTagNodes(CAtlArray<CXNodeIndex>& basenodes, CAtlArray<CXNodeIndex>& nodes, int rate, int nCount)
{
	int i;
	POSITION pos;
	CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair* pPair;
	CXTree* pTree;

	for(i = 0; i < 26; i ++)
		if(pTree = m_aSubTrees[i])
			pTree->FillTagNodes(basenodes, nodes, rate, nCount);

	pos = m_aTagNodes.GetHeadPosition();
	while(pos)
	{
		pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.GetNext(pos);

		for(i = 0; i < (int)basenodes.GetCount(); i ++)
			if(!pPair->m_key.Compare(basenodes[i].m_pNode->m_strTags[0]))
				break;

		if(i == basenodes.GetCount())
		{
			CXNodeIndex node;
			POSITION pos;

			node.m_pNode = NULL;
			pos = pPair->m_value->GetHeadPosition();
			while(pos)
			{
				node.m_pNode = pPair->m_value->GetNext(pos);
				if(node.m_pNode->m_nLife > 0)
					break;
				node.m_pNode = NULL;
			}

			if(node.m_pNode)
			{
				node.m_nLife = node.m_pNode->m_nLife * rate;

				for(i = 0; i < (int)nodes.GetCount(); i ++)
					if(node.m_nLife >= nodes[i].m_nLife || (!pPair->m_key.Compare(nodes[i].m_pNode->m_strTags[0])))
						break;

				if((i == (int)nodes.GetCount()) || node.m_nLife >= nodes[i].m_nLife)
				{
					if(i < nCount)
						nodes.InsertAt(i, node);

					i ++;
					while(i < (int)nodes.GetCount())
						if((!pPair->m_key.Compare(nodes[i].m_pNode->m_strTags[0])))
							nodes.RemoveAt(i);
						else i ++;

					if((int)nodes.GetCount() > nCount)
						nodes.RemoveAt(nCount, nodes.GetCount() - nCount);
				}
			}
		}
	}
}

void CXGeoLinks::CXTree::GetTagNodes(LPCWSTR mark, CAtlArray<CXNodeIndex>& basenodes, CAtlArray<CXNodeIndex>& nodes, int nCount, int *rates)
{
	CXTree* pTree;
	LPCWSTR mark1;
	int i;

	for(i = 0; i < 26; i ++)
		if(rates[i] && (pTree = m_aSubTrees[i]))
		{
			mark1 = mark;
			WCHAR ch = *mark1;

			while(ch)
			{
				mark1 ++;
				ch = (ch & 0x1f) - 1;

				if(!pTree->m_aSubTrees[ch])
				{
					pTree = NULL;
					break;
				}

				pTree = pTree->m_aSubTrees[ch];
				ch = *mark1;
			}

			if(pTree)
				pTree->FillTagNodes(basenodes, nodes, rates[i], nCount);
		}
}

void CXGeoLinks::CXRoot::GetTagNodes(int cityid, LPCWSTR mark, int nCount, int *rates, CAtlArray<CXNodeIndex>& nodes)
{
	CAtlArray<CXNodeIndex> nodes1;
	int nCount1 = nCount;
	int p1, p2;
	int cityid1;

	CityInfoStruct* pCity = (CityInfoStruct*)bsearch(&cityid, s_cityname, CITYCOUNT, sizeof(s_cityname[0]), _cmpcityid);

	if(pCity == NULL)
		return;

	p1 = p2 = pCity - s_cityname;
	cityid1 = nextcity(cityid);

	while(nCount1 > 0 && ((p1 >= 0 && s_cityname[p1].id > cityid) || (p2 < CITYCOUNT && s_cityname[p2].id < cityid1)))
	{
		while(p1 >= 0 && s_cityname[p1].id > cityid)
		{
			p1 --;

			if(p1 >= 0 && m_mapCitys[p1])
				m_mapCitys[p1]->GetTagNodes(mark, nodes, nodes1, nCount1, rates);
		}

		while(p2 < CITYCOUNT && s_cityname[p2].id < cityid1)
		{
			if(p2 < CITYCOUNT && m_mapCitys[p2])
				m_mapCitys[p2]->GetTagNodes(mark, nodes, nodes1, nCount1, rates);

			p2 ++;
		}

		nodes.Append(nodes1);
		nodes1.RemoveAll();

		nCount1 = nCount - nodes.GetCount();

		cityid = parentcity(cityid);
		cityid1 = nextcity(cityid);
	}
}

STDMETHODIMP CXGeoLinks::GetTagList(int cityid, BSTR mark, int nStart, int nCount, BSTR Types, VARIANT* pVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	CAtlArray<CXNodeIndex> nodes;
	int rates[26];
	int i, l, n;
	CXComPtr<CXRoot> root;

	n = 0;
	l = ::SysStringLen(Types);
	for(i = 0; i < 26; i ++)
	{
		if(i >= l)
			rates[i] = 0;
		else n += rates[i] = (Types[i] & 0x1f) - 1;
	}

	if(n > 0)
	{
		n = ::SysStringLen(mark);
		if(n > 15)return E_INVALIDARG;

		for(i = 0; i < n; i ++)
			if(!((mark[i] >= 'A' && mark[i] <= 'Z') || (mark[i] >= 'a' && mark[i] <= 'z')))
				return E_INVALIDARG;

		root = m_root;
		root->GetTagNodes(cityid, mark, nStart + nCount, rates, nodes);
	}

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	CXNode* pNode;

	if((int)nodes.GetCount() > nStart)
	{
		hr = bstrArray.Create(nodes.GetCount() - nStart);
		if(FAILED(hr))return hr;

		pVar = (VARIANT*)bstrArray.m_psa->pvData;

		for(i = nStart; i < (int)nodes.GetCount(); i ++)
		{
			pNode = nodes[i].m_pNode;
			pVar->vt = VT_BSTR;
			pVar->bstrVal = ::SysAllocStringLen(NULL, pNode->m_strTags[0].GetLength() + pNode->m_strValue.GetLength() + 1);
			CopyMemory(pVar->bstrVal, pNode->m_strTags[0], pNode->m_strTags[0].GetLength() * 2);
			pVar->bstrVal[pNode->m_strTags[0].GetLength()] = '\t';
			CopyMemory(pVar->bstrVal + pNode->m_strTags[0].GetLength() + 1, pNode->m_strValue, pNode->m_strValue.GetLength() * 2);
			pVar ++;
		}
	}else
	{
		hr = bstrArray.Create();
		if(FAILED(hr))return hr;
	}

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

void CXGeoLinks::CXTree::FillHotTags(CAtlList< CXComPtr<CXNode> >* pList, LPCWSTR tag, CRBMap<CXString, int>& nodes, double& n)
{
	CRBMap<CXString, int>::CPair* pPair1;
	POSITION pos;
	CXNode* pnode;
	int i;

	pos = pList->GetHeadPosition();

	while(pos)
	{
		pnode = pList->GetNext(pos);

		for(i = 0; i < (int)pnode->m_strTags.GetCount(); i ++)
			if(pnode->m_strTags[i].CompareNoCase(tag))
			{
				pPair1 = nodes.Lookup(pnode->m_strTags[i]);
				
				if(pPair1)
				{
					pPair1->m_value ++;
					if(pPair1->m_value > n)
						n = pPair1->m_value;
				}else
				{
					nodes.SetAt(pnode->m_strTags[i], 1);
					if(1 > n)
						n = 1;
				}
			}
	}
}

void CXGeoLinks::CXTree::CountHotTags(LPCWSTR tag, CRBMap<CXString, int>& nodes, double& n)
{
	int i;
	CXTree* pTree;
	CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair* pPair;

	if(tag && *tag)
	{
		if(pPair = m_aTagNodes.Lookup(tag))
			FillHotTags(pPair->m_value, tag, nodes, n);
		if(pPair = m_aTagNodes1.Lookup(tag))
			FillHotTags(pPair->m_value, tag, nodes, n);
	}else
	{
		POSITION pos;
		CRBMap<CXString, int>::CPair* pPair1;

		pos = m_aTagNodes.GetHeadPosition();
		while(pos)
		{
			pPair = (CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* >::CPair*)m_aTagNodes.GetNext(pos);

			pPair1 = nodes.Lookup(pPair->m_key);

			if(pPair1)
			{
				pPair1->m_value += pPair->m_value->GetCount();
				if(pPair1->m_value > n)
					n = pPair1->m_value;
			}else
			{
				i = pPair->m_value->GetCount();
				nodes.SetAt(pPair->m_key, i);
				if(i > n)
					n = i;
			}
		}
	}

	for(i = 0; i < 26; i ++)
		if(pTree = m_aSubTrees[i])
			pTree->CountHotTags(tag, nodes, n);
}

void CXGeoLinks::CXTree::GetHotTags(LPCWSTR mark, LPCWSTR tag, CRBMap<CXString, int>& nodes, double& n)
{
	CXTree* pTree;
	LPCWSTR mark1;
	int i;

	for(i = 0; i < 26; i ++)
		if(pTree = m_aSubTrees[i])
		{
			mark1 = mark;
			WCHAR ch = *mark1;

			while(ch)
			{
				mark1 ++;
				ch = (ch & 0x1f) - 1;

				if(!pTree->m_aSubTrees[ch])
				{
					pTree = NULL;
					break;
				}

				pTree = pTree->m_aSubTrees[ch];
				ch = *mark1;
			}

			if(pTree)
				pTree->CountHotTags(tag, nodes, n);
		}
}

void CXGeoLinks::CXRoot::GetHotTags(int cityid, LPCWSTR mark, LPCWSTR tag, int nCount, CAtlArray<CXTagIndex>& nodes)
{
	int p1, p2, i;
	int cityid1;
	double htValue = 9;
	double n, n1;
	CRBMap<CXString, int> nodes1;
	CRBMap<CXString, int>::CPair* pPair;
	POSITION pos;

	CityInfoStruct* pCity = (CityInfoStruct*)bsearch(&cityid, s_cityname, CITYCOUNT, sizeof(s_cityname[0]), _cmpcityid);

	if(pCity == NULL)
		return;

	p1 = p2 = pCity - s_cityname;
	cityid1 = nextcity(cityid);

	while((p1 >= 0 && s_cityname[p1].id > cityid) || (p2 < CITYCOUNT && s_cityname[p2].id < cityid1))
	{
		n = 0;
		while(p1 >= 0 && s_cityname[p1].id > cityid)
		{
			p1 --;

			if(p1 >= 0 && m_mapCitys[p1])
				m_mapCitys[p1]->GetHotTags(mark, tag, nodes1, n);
		}

		while(p2 < CITYCOUNT && s_cityname[p2].id < cityid1)
		{
			if(p2 < CITYCOUNT && m_mapCitys[p2])
				m_mapCitys[p2]->GetHotTags(mark, tag, nodes1, n);

			p2 ++;
		}

		n = htValue / n;

		pos = nodes1.GetHeadPosition();
		while(pos)
		{
			pPair = (CRBMap<CXString, int>::CPair*)nodes1.GetNext(pos);

			n1 = pPair->m_value * n;
			for(i = 0; i < (int)nodes.GetCount(); i ++)
				if(nodes[i].m_nLife < n1)
				{
					CXTagIndex node;

					node.m_strTag = pPair->m_key;
					node.m_nLife = n1;
					nodes.InsertAt(i, node);

					for(i ++; i < (int)nodes.GetCount(); i ++)
						if(!pPair->m_key.CompareNoCase(nodes[i].m_strTag))
						{
							nodes.RemoveAt(i);
							break;
						}

					if((int)nodes.GetCount() > nCount)
						nodes.RemoveAt(nCount, nodes.GetCount() - nCount);

					i = nCount;
					break;
				}else if(!pPair->m_key.CompareNoCase(nodes[i].m_strTag))
				{
					i = nCount;
					break;
				}

			if(i < nCount && i == (int)nodes.GetCount())
			{
				CXTagIndex node;

				node.m_strTag = pPair->m_key;
				node.m_nLife = n1;
				nodes.Add(node);
			}
		}

		nodes1.RemoveAll();

		htValue /= 2;

		cityid = parentcity(cityid);
		cityid1 = nextcity(cityid);
	}
}

STDMETHODIMP CXGeoLinks::GetHotTag(int cityid, BSTR mark, BSTR tag, int nStart, int nCount, VARIANT* pVal)
{
	if(m_pUpdate)
		return DISP_E_BADCALLEE;

	CAtlArray<CXTagIndex> nodes;
	int i;
	CXComPtr<CXRoot> root;

	root = m_root;
	root->GetHotTags(cityid, mark, tag, nStart + nCount, nodes);

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;

	if((int)nodes.GetCount() > nStart)
	{
		hr = bstrArray.Create(nodes.GetCount() - nStart);
		if(FAILED(hr))return hr;

		pVar = (VARIANT*)bstrArray.m_psa->pvData;

		for(i = nStart; i < (int)nodes.GetCount(); i ++)
		{
			pVar->vt = VT_BSTR;
			pVar->bstrVal = ::SysAllocStringLen(NULL, nodes[i].m_strTag.GetLength() + 2);
			pVar->bstrVal[0] = (int)(sqrt(nodes[i].m_nLife) * 3) + '0';
			pVar->bstrVal[1] = '\t';
			CopyMemory(pVar->bstrVal + 2, nodes[i].m_strTag, nodes[i].m_strTag.GetLength() * 2);
			pVar ++;
		}
	}else
	{
		hr = bstrArray.Create();
		if(FAILED(hr))return hr;
	}

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}
