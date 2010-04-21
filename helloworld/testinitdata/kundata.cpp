#include "kundata.h"
#include <time.h>
#include <math.h>
#include <cstdlib>
#include <iostream> 
using namespace std;
kundata::kundata(void)
{
	initdata();
	initprobability();
}

kundata::~kundata(void)
{
}

int kundata::initprobability()
{
	srand((unsigned)time(NULL));
	//参与查询的数据
	for(int i=0; i < MAXN; ++i) {
		kdata[i] =  rand() % 1000 / 1000.0;
	}
	return 1;
}

int kundata::initdata()
{
	//srand()函数产生一个以当前时间开始的随机种子.应该放在for等循环语句前面 不然要很长时间等待
	srand((unsigned)time(NULL));
	//参与查询的数据
	for(int i=0; i < MAXN; ++i) {
		kdata[i] =  rand() % MAXN;
	}
	return 1;
}

kdataitem & kundata::operator[]( const size_t index)
{
	return kdata[index];
}

const kdataitem & kundata::operator[]( const size_t index) const
{
	return kdata[index];
}

int* kundata::readdatalist()
{
	int ndatalist[MAXN];
	for (int i = 0; i < MAXN; i++)
	{
		ndatalist[i] = kdata[i].readdata();
	}
	return ndatalist;
}

double* kundata::readprobabilitylist()
{
	double dprobabilitylist[MAXN];
	for (int i = 0; i < MAXN; i++)
	{
		dprobabilitylist[i] = kdata[i].readprobability();
	}
	return dprobabilitylist;
}
std::ostream & operator <<(std::ostream &out,kundata &item)
{ 
	for (int i = 0;i < MAXN; i++)
	{
		out << item.kdata[i];
	}
	return out;
}

std::ostream & operator <<(std::ostream &os,kdataitem &item )
{
	os << item.m_data  <<" : " << item.m_probability <<endl; 
	return os;
}