#pragma once
#include "kundata.h"
#include <cstring>
using namespace std;
class kreorderlist
{
public:
	kreorderlist(){}
	~kreorderlist(){}
	int readgroupcount(){return m_ngroupcount;}
	int readitemcount(){return m_nitemcount;}
	int* readgroup(int);
	int putgroup(int*);
	int movegroup(int*);	
private:
	static const char itemsplit = ',';
	static const char groupsplit = ';';
	int m_ngroupcount;
	int m_nitemcount;
	int m_ngroupserial;
	int m_nitemserial;
	CString m_strList;
};
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
	int compression(int);
	void subsetprovalue(int);
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
