#pragma once
#include "hcode.h"
#include <iostream> 
using namespace std;
class kdataitem
{
public:
	kdataitem(){m_mark = false;}
	~kdataitem(){}
	friend std::ostream & operator <<(std::ostream &,kdataitem &); 
	//friend std::ostream & operator <<(std::ostream&,const kdataitem&);
	//void setdata(int &ndata)const{m_data = ndata;};
	//void setprobability(double &probability)const{m_probability = probability;};
	double readprobability(){return m_probability;}
	int readdata(){return m_data;}
	void setmark(){m_mark = true;}
	bool readmark(){return m_mark;}
	kdataitem& operator=(const double &dprobability){m_probability = dprobability;return *this;}
	kdataitem& operator=(const int &ndata){m_data = ndata;return *this;}
private:
	double m_probability;
	int m_data;
	bool m_mark;
};
class kundatabase
{
public:
	kundatabase(void)
	{
		m_nsize = MAXN;
		init();
	}
	kundatabase(const size_t nsize)
	{
		m_nsize = nsize;init();
	}
	virtual int initprobability() = ZERO;
	virtual int initdata() = ZERO;
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
	int makerule(kdataitem*,int);
	void display(double*);
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
	kundata(void);
	int initprobability();
	int initdata();
	int initrule();
	int* readdatalist();
	double* readprobabilitylist();
	friend std::ostream & operator <<(std::ostream&,kundata&);
	kdataitem &operator[] (const size_t);
	const kdataitem &operator[] (const size_t)const;
	~kundata(void);
private:
	kruleitem* m_pkritem;
};





