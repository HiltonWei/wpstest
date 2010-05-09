#include "ktopquery.h"
#include <cstring>
using namespace std;
ktopquery* ktopquery::instance = NULL;
ktopquery::ktopquery(void)
{
	init();
}


ktopquery::~ktopquery(void)
{
}

int ktopquery::init()
{
	count_sw = 0;
	count_cmp = 0;
	m_pkdata = kundata::GetSingleton();
	m_pdatalist = m_pkdata->readdatalist();
	m_pprolist = m_pkdata->readprobabilitylist();
	m_pkitemlist = m_pkdata->readitemlist();
	arr = new int[koption::s_nMaxnum];
	for (int i = 0; i < koption::s_nMaxnum; i++)
	{
		arr[i] = i;
	}
	topK(0, koption::s_nMaxnum-1);
	for (int i = 0; i < koption::s_nMaxnum; i++)
	{
		compression(i);
		subsetprovalue(i);
	}
	
	return 1;
}

void ktopquery::swap( int& x,int& y)
{
	int t ;
	t = x;
	x = y;
	y = t;
	return ;
}

void ktopquery::qsort(int start, int end )
{
	if(start < end) {
		int mid = arr[rand()%(end-start) + start];
		int i = start - 1;
		int j = end + 1;
		while (true)
		{
			while ((m_pkitemlist[arr[++i]]).readdata() > (m_pkitemlist[mid]).readdata()) count_cmp++;
			while ((m_pkitemlist[arr[--j]]).readdata() < (m_pkitemlist[mid]).readdata()) count_cmp++;
			if(i>=j) break;
			swap(arr[i], arr[j]);
			count_sw++;
		}
		qsort (start, i-1);
		qsort (j+1, end);
	}
	return ;
}

void ktopquery::topK(int start, int end)
{
	int k = koption::s_nK;
	if(start < end) {
		int mid = arr[rand()%(end-start) + start];
		int i = start - 1;
		int j = end + 1;
		while (true) {
			while (m_pkitemlist[arr[++i]].readdata() > m_pkitemlist[mid].readdata()) count_cmp++;
			while (m_pkitemlist[arr[--j]].readdata() < m_pkitemlist[mid].readdata()) count_cmp++;
			if(i>=j) break;
			swap(arr[i], arr[j]);
			count_sw++;
		}
		if(i-start > k) {
			last = i-1;
			topK (start, i-1);
		} else{
			qsort(start, last);
		}
	}
	return ;
}

int ktopquery::compression(int i)
{
	if (i = 1)return i;

	if (!m_pkitemlist[arr[i]].readmark())return -1;
	if ()
	{
	}


}

void ktopquery::subsetprovalue(int i)
{

}
std::ostream & operator <<(std::ostream &os,ktopquery &item)
{
	for(int i = 0; i < koption::s_nK; i++) 
	{
		os << item.m_pkitemlist[item.arr[i]];
	}
	os << "Swaps:" << item.count_sw << endl;
	os << "Campare:" << item.count_cmp << endl;
	return os;
}

int* kreorderlist::readgroup( int i)
{

}