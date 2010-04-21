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

int main() 
{
	using std::cout; 
	using std::endl; 
	record a; 
	cout <<a; 
	return 0;
} 

