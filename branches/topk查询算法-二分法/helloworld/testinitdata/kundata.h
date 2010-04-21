#pragma once
#include "hcode.h"
#include <iostream> 
using namespace std;
class kdataitem
{
public:
	kdataitem(){}
	~kdataitem(){}
	friend std::ostream & operator <<(std::ostream &,kdataitem &); 
	//friend std::ostream & operator <<(std::ostream&,const kdataitem&);
	//void setdata(int &ndata)const{m_data = ndata;};
	//void setprobability(double &probability)const{m_probability = probability;};
	double readprobability(){return m_probability;}
	int readdata(){return m_data;}	
	kdataitem& operator=(const double &dprobability){m_probability = dprobability;return *this;}
	kdataitem& operator=(const int &ndata){m_data = ndata;return *this;}
private:
	double m_probability;
	int m_data;
};
class kundatabase
{
public:
	friend kdataitem;
	kundatabase(void){}
	virtual int initprobability() = 0;
	virtual int initdata() = 0;
	~kundatabase(void){}
	kdataitem kdata[MAXN];

	
};
class kundata :
	virtual public kundatabase
{
public:
	kundata(void);
	int initprobability();
	int initdata();
	int* readdatalist();
	double* readprobabilitylist();
	friend std::ostream & operator <<(std::ostream&,kundata&);
	kdataitem &operator[] (const size_t);
	const kdataitem &operator[] (const size_t)const;
	~kundata(void);
};
// std::ostream & operator <<(std::ostream& out,const kdataitem& item)
// {
// 	out << item.m_data << "  " << item.m_probability;
// 	return out;
// }




