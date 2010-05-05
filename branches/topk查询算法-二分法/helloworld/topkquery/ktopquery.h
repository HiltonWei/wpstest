#pragma once
#include "kundata.h"
class ktopquery
{
public:
	ktopquery(void);
	~ktopquery(void);
	int init();
	void swap(int&,int&);
	void qsort(int start, int end);
	void topK(int start, int end, int k);
private:
	int last;
	int count_sw, count_cmp;
	int* m_pdatalist;
	double* m_pprolist;
	kdataitem* m_pkitemlist;
	kundata* m_pkdata;
};
