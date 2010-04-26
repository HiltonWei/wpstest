#include "kundata.h"
#include <time.h>
#include <math.h>
#include <cstdlib>
#include <iostream> 
#include <Windows.h>
using namespace std;
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
		m_pkdata[i] =  rand() % 1000 / 1000.0;
	}
	return 1;
}

int kundata::initdata()
{
	//srand()函数产生一个以当前时间开始的随机种子.应该放在for等循环语句前面 不然要很长时间等待
	srand((unsigned)time(NULL) + rand());
	//参与查询的数据
	for(int i=0; i < this->readsize(); ++i) {
		m_pkdata[i] =  rand() % this->readsize();
	}
	return 1;
}

kdataitem & kundata::operator[]( const size_t index)
{
	return m_pkdata[index];
}

const kdataitem & kundata::operator[]( const size_t index) const
{
	return m_pkdata[index];
}

int* kundata::readdatalist()
{
	int* pndatalist = new int[this->readsize()];
	for (int i = ZERO; i < this->readsize(); i++)
	{
		pndatalist[i] = m_pkdata[i].readdata();
	}
	return pndatalist;
}

double* kundata::readprobabilitylist()
{
	double* pdprobabilitylist = new double[this->readsize()];
	for (int i = ZERO; i < this->readsize(); i++)
	{
		pdprobabilitylist[i] = m_pkdata[i].readprobability();
	}
	return pdprobabilitylist;
}

int kundata::initrule()
{
	m_pkritem = new kruleitem[MULTIRULE];
	for (int i = ZERO; i < MULTIRULE; i++)
	{
		m_pkritem[i].makerule(m_pkdata, readsize());
	}
	return 1;
}
std::ostream & operator <<(std::ostream &out,kundata &item)
{ 
	for (int i = ZERO;i < item.readsize(); i++)
	{
		out << item.m_pkdata[i];
	}
	return out;
}

std::ostream & operator <<(std::ostream &os,kdataitem &item )
{
	os << item.m_data  <<" : " << item.m_probability <<endl; 
	return os;
}

kruleitem::kruleitem(void)
{
	m_nrulesizemin = RULEMIN;
	m_nrulesizemax = RULEMAX;
	initruleitem();
}

int kruleitem::initruleitem()
{
	int nsplit;
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

int kruleitem::makerule(kdataitem* pklist, int nsize)
{
	int nlistsize = nsize;
	int nord;	
	srand((unsigned)time(NULL) + rand());
	cout<<"rule begin"<<endl;
	for (int i = 0; i < m_nrulecount; i++)
	{
		do 
		{
			m_pnordlist[i] = rand() % (nlistsize - 1);
		} while (pklist[m_pnordlist[i]].readmark());	
		pklist[m_pnordlist[i]].setmark();
		pklist[m_pnordlist[i]] = m_pdruleprolist[i];
		cout<<pklist[m_pnordlist[i]];
	}
	cout<<"rule end"<<endl;
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