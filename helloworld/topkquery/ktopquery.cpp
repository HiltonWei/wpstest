#include "ktopquery.h"
using namespace std;
ktopquery::ktopquery(void)
{
	init();
}

ktopquery::~ktopquery(void)
{
}

int ktopquery::init()
{
	m_pkdata = kundata::GetSingleton();
	m_pdatalist = m_pkdata->readdatalist();
	m_pprolist = m_pkdata->readprobabilitylist();
	m_pkitemlist = m_pkdata->readitemlist();
	return 1;
}

void ktopquery::swap( kdataitem& x,kdataitem& y)
{
	kdataitem t ;
	t = x;
	x = y;
	y = t;
	return ;
}

void ktopquery::qsort(int start, int end )
{

}

void ktopquery::topK(int start, int end, int k )
{

}