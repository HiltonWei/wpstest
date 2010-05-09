#include "kundata.h"
#include <time.h>
#include <math.h>
#include <cstdlib>
#include <iostream> 
#include <fstream>
#include <Windows.h>
using namespace std;
int koption::s_nMaxnum = MAXN;
int koption::s_nLimit = LIMIT;
int koption::s_nMultirule = MULTIRULE;
int koption::s_nRulemax = RULEMAX;
int koption::s_nRulemin = RULEMIN;
int koption::s_nK = KVALUE;
kundata* kundata::instance = NULL;
koption* koption::instance = NULL;
kundata::kundata(void)
{
	initdata();
	initprobability();
	initrule();
}

kundata::~kundata(void)
{
}

int kundata::initprobability()
{
	srand((unsigned)time(NULL) + rand());
	//参与查询的数据
	for(int i=ZERO; i < this->readsize(); ++i) {
		m_pkdatalist[i] =  rand() % 1000 / 1000.0;
	}
	return 1;
}

int kundata::initdata()
{
	const char filename[] = MYRULEFILE;
	//srand()函数产生一个以当前时间开始的随机种子.应该放在for等循环语句前面 不然要很长时间等待
	srand((unsigned)time(NULL) + rand());
	//参与查询的数据
	for(int i=0; i < this->readsize(); ++i) {
		m_pkdatalist[i] =  rand() % this->readsize();
	}
	return 1;
}

kdataitem & kundata::operator[]( const size_t index)
{
	return m_pkdatalist[index];
}

const kdataitem & kundata::operator[]( const size_t index) const
{
	return m_pkdatalist[index];
}

int* kundata::readdatalist()
{
	int* pndatalist = new int[this->readsize()];
	for (int i = ZERO; i < this->readsize(); i++)
	{
		pndatalist[i] = m_pkdatalist[i].readdata();
	}
	return pndatalist;
}

double* kundata::readprobabilitylist()
{
	double* pdprobabilitylist = new double[this->readsize()];
	for (int i = ZERO; i < this->readsize(); i++)
	{
		pdprobabilitylist[i] = m_pkdatalist[i].readprobability();
	}
	return pdprobabilitylist;
}

int kundata::initrule()
{
	const char filename[] = MYRULEFILE;
	ofstream o_file;
	o_file.open(filename);
	o_file << *(koption::GetSingleton());	
	m_pkritem = new kruleitem[koption::s_nMultirule];
	for (int i = ZERO; i < koption::s_nMultirule; i++)
	{
		m_pkritem[i].makerule(o_file,this->m_pkdatalist,i);
	}	
	o_file << *this;
	o_file.close();
	return 1;
}
std::ostream & operator <<(std::ostream &out,kundata &item)
{ 
	for (int i = ZERO;i < item.readsize(); i++)
	{
		out << item[i];
	}
	return out;
}

std::ostream & operator <<(std::ostream &os,kdataitem &item )
{
	os << item.m_data  <<" : " 
		<< item.m_probability <<" : " 
		<< item.m_ruleserial
		<<endl; 
	return os;
}

std::ostream & operator <<(std::ostream &os,koption &item )
{
	os << "Data number: " << koption::s_nMaxnum
		<< "  Rule number: " << koption::s_nMultirule
		<< "  Rule float: " << koption::s_nRulemin << " - " 
		<< koption::s_nRulemax << endl; 
	return os;
}
std::ostream & operator <<(std::ostream & os,kruleitem & item)
{
	for (int i = 0; i < item.m_nrulecount; i++)
	{
		os<< item.m_pnordlist[i] <<" : " << item.m_pdruleprolist[i] << endl;
	}
	os << endl;
	return os;
}
kruleitem::kruleitem(void)
{//初始化一个rule
	m_nrulesizemin = koption::s_nRulemin;
	m_nrulesizemax = koption::s_nRulemax;
	initruleitem();
}

int kruleitem::initruleitem()
{//获取一个generation rule
	int nsplit;//分隔符
	srand((unsigned)time(NULL) + rand());
	m_nrulecount = m_nrulesizemin + rand() % (m_nrulesizemax - m_nrulesizemin + 1);
	nsplit = m_nrulecount - 1;
	double* pdsplitlist = new double[m_nrulecount];
	m_pdruleprolist = new double[m_nrulecount];
	m_pnordlist = new int[m_nrulecount];

	pdsplitlist[nsplit] = 1.0;
	for (int i = 0; i < nsplit; i++)
	{
		pdsplitlist[i] =  rand() % 1000 / 1000.0;
	}

	for(int i=0; i < nsplit; i++)
	{
		for(int j = nsplit-1; j >= i; j--)
		{
			if(pdsplitlist[j] < pdsplitlist[j-1])
			{
				double temp = pdsplitlist[j];
				pdsplitlist[j] = pdsplitlist[j-1];
				pdsplitlist[j-1] = temp;
			}
		}
	}
	m_pdruleprolist[0] = pdsplitlist[0];
	//display(pdsplitlist);
	for(int i = m_nrulecount - 1; i > 0; i--)
	{
		m_pdruleprolist[i] = pdsplitlist[i] - pdsplitlist[i - 1];
	}
	srand((unsigned)time(NULL) + rand());
	//display(m_pdruleprolist);
	for(int i = 0; i < m_nrulecount; i++)
	{
		int temp0 = rand() % 2;
		double temp1 = m_pdruleprolist[i] - rand() % 1000 / 1000.0;
		if (temp0 == 1 && temp1 > 0)
		{
			m_pdruleprolist[i] = temp1;
		}		
	}
	return 1;
}

int kruleitem::makerule(std::ostream &out,kdataitem* pklist,int nserial)
{
	int nlistsize = koption::s_nMaxnum;
	int nord;
	out << nserial<<" rule begin\n";
	for (int i = 0; i < m_nrulecount; i++)
	{
		do 
		{
			m_pnordlist[i] = rand() % (nlistsize - 1);
		} while (pklist[m_pnordlist[i]].readmark());	
		pklist[m_pnordlist[i]].setmark();//设置标志位
		pklist[m_pnordlist[i]] = m_pdruleprolist[i];//修改概率使其满足规则
		(pklist[m_pnordlist[i]]).setruleserial(nserial);//设置规则序号
		out << "  " << pklist[m_pnordlist[i]];		
	}
	out << "rule end\n";
	return 1;
}

void kruleitem::display(double* plist)
{
	for (int i = 0; i < m_nrulecount; i++)
	{
		cout<<plist[i]<<"  ";
	}
	cout<<endl;
	return;
}