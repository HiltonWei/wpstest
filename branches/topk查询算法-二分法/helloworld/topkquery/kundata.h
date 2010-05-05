#pragma once
#include "hcode.h"
#include <iostream> 
using namespace std;
class koption
{
public:
	static koption *GetSingleton()
	{
		static bool inited = false;
		if (!inited)
		{
			instance = new koption();
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
	static int s_nMultirule;//rule个数
	static int s_nRulemin;//一个rule最小个数
	static int s_nRulemax;//一个rule最大个数
	static int s_nMaxnum;//元组的个数
	static int s_nLimit;//保留参数
	static int s_nK;
	friend std::ostream & operator <<(std::ostream&,koption&);
private:
	static koption *instance;
	koption(void){}
	~koption(void){}
};


class kdataitem
{
public:
	kdataitem(){m_mark = false;m_ruleserial = -1;}
	~kdataitem(){}
	friend std::ostream & operator <<(std::ostream &,kdataitem &); 
	double readprobability(){return m_probability;}
	int readdata(){return m_data;}
	void setmark(){m_mark = true;}
	bool readmark(){return m_mark;}
	int readruleserial(){return m_ruleserial;}
	void setruleserial(int number){m_ruleserial = number;}
	kdataitem& operator=(const double &dprobability){m_probability = dprobability;return *this;}
	kdataitem& operator=(const int &ndata){m_data = ndata;return *this;}
private:
	double m_probability;
	int m_data;
	bool m_mark;
	int m_ruleserial;
};
class kundatabase
{
public:
	kundatabase(void)
	{
		m_nsize = koption::s_nMaxnum;
		init();
	}
	virtual int initprobability() = ZERO;
	virtual int initdata() = ZERO;
	kdataitem* readitemlist(){return m_pkdatalist;}
	int readsize(){return m_nsize;}
	~kundatabase(void){}
	kdataitem* m_pkdatalist;	
private:
	int m_nsize;
	void init()
	{
		m_pkdatalist = new kdataitem[m_nsize];
	}
};

class kruleitem
{
public:
	kruleitem(void);
	~kruleitem(void){};
	int initruleitem();
	int makerule(std::ostream &,kdataitem*,int);
	void display(double*);
	friend std::ostream & operator <<(std::ostream &,kruleitem &); 
protected:
	int m_nrulesizemin;
	int m_nrulesizemax;
	int m_nrulecount;
	int m_nordnum;
	double* m_pdruleprolist;
	int* m_pnordlist;
};
class kundata :
	virtual public kundatabase
{
public:
	static kundata *GetSingleton()
	{
		static bool inited = false;
		if (!inited)
		{
			instance = new kundata();
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
	int initprobability();
	int initdata();
	int initrule();
	int* readdatalist();
	double* readprobabilitylist();
	friend std::ostream & operator <<(std::ostream&,kundata&);
	kdataitem &operator[] (const size_t);
	const kdataitem &operator[] (const size_t)const;
	
private:
	static kundata *instance;
	kundata(void);
	~kundata(void);
	kruleitem* m_pkritem;
};





