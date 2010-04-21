#include "kundata.h"
#include <time.h>
#include <math.h>
#include <cstdlib>

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
	//�����ѯ������
	for(int i=0; i < MAXN; ++i) {
		kdata[i] =  rand() % 1000 / 1000.0;
	}
	return 1;
}

int kundata::initdata()
{
	//srand()��������һ���Ե�ǰʱ�俪ʼ���������.Ӧ�÷���for��ѭ�����ǰ�� ��ȻҪ�ܳ�ʱ��ȴ�
	srand((unsigned)time(NULL));
	//�����ѯ������
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


std::ostream & operator <<(std::ostream& out,const kundata& item)
{
	using std::endl; 
	for (int i = 0;i < 10; i++)
	{
		out << item.kdata[i].readdata() << "  " << item.kdata[i].readprobability()<< endl;
	}
	return out;
}
