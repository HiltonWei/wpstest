#pragma once
#include "kundata.h"
#include <iostream> 
#include <string>
#include <sstream>
using namespace std;
class kreorderlist
{
public:
	static kreorderlist *GetSingleton()
	{
		static bool inited = false;
		if (!inited)
		{
			instance = new kreorderlist();
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
	/* States */
	enum States {Removed, Empty, Added}; // states
	friend std::ostream & operator <<(std::ostream&,kreorderlist&);
	
	int readgroupcount(){return m_ngroupcount;}
	int* readgroup(int);
	int putgroup(int*);
	int enterRemove(int,int);
	int enterAdd(int,int);
	int enterJoin(int, int);
	bool checkgroup(int);
	void putitem(int,int);
	void changeState(States);
	States currentState() { return m_state; }
private:
	static kreorderlist *instance;
	kreorderlist()
	{
		m_ngroupcount = 0;
		m_state = Empty;
	}
	~kreorderlist(){}
	static const char itemsplit = ',';
	static const char groupsplit = ' ';
	States m_state;
	int m_ngroupcount;//total group
	int m_ngroupserial;//group serial in strlist
	int m_nremovegroupnum;//remove item num
	string m_groupnumlist;//group numnber
	string m_strList;//data serial
	string m_strremovegroup;//store the remove data
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
