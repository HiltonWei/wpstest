// test.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream> 
using namespace std;
class time 
{
private: 
	int h; 
	int m; 
	int s; 
public: 
	time(int a,int b,int c); 
	time(); 
	friend std::ostream & operator <<(std::ostream &os,time &t); 
}; 
class record
{
public:
	//friend time;
	record(){m_a = new time(3,45,48);}
	~record(){}	
	friend std::ostream & operator <<(std::ostream &os,record &t); 
private:
	time *m_a;
};


time::time(int a,int b,int c) 
{
	h=a; 
	m=b; 
	s=c;} 

time::time() 
{
	h=m=s=0; 
} 

std::ostream & operator <<(std::ostream &os,time &t) 
{
	using std::endl; 
	os <<t.h <<":" <<t.m <<":" <<t.s <<endl; 
	return os;
} 

std::ostream & operator<<( std::ostream &os,record &t )
{
	os << *(t.m_a);
	return os;
}

void dispaly(double* list)
{
	for (int i=0; i <10; i++)
	{
		cout<<list[i]<<"   ";
	}
	cout<<endl;
}
int main() 
{
	int m_nrulecount = 10;
	double m_pruleprolist[10];
	for (int i = 9; i >= 0; i--)
	{
		m_pruleprolist[i] =  10-i;
	}
	for(int i=0; i < m_nrulecount; i++)
	{
		dispaly(m_pruleprolist);
		for(int j = m_nrulecount-1; j >= i; j--)
		{
			if(m_pruleprolist[j] < m_pruleprolist[j-1])
			{
				double temp = m_pruleprolist[j];
				m_pruleprolist[j] = m_pruleprolist[j-1];
				m_pruleprolist[j-1] = temp;
			}			
		}
		
	}
	system("PAUSE");
	return 0;
} 

