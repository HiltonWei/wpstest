#pragma once
#include "kundata.h"
class ktopquery
{
public:
	static ktopquery *GetSingleton()
	{
		static bool inited = false;
		if (!inited)
		{
			instance = new ktopquery();
			inited = true;
		}
		return instance;        
	}
	static void Release()
	{
		if (NULL != instance)
		{
			delete instance;
			instance = NULL;
		}
	}
	
	int init();
	void swap(int&,int&);
	void qsort(int start, int end);
	void topK(int start, int end);
	friend std::ostream & operator <<(std::ostream &,ktopquery &); 
private:
	int last;
	int* arr;
	int count_sw, count_cmp;
	int* m_pdatalist;
	double* m_pprolist;
	static ktopquery *instance;
	ktopquery(void);
	~ktopquery(void);
	kdataitem* m_pkitemlist;
	kundata* m_pkdata;
};
