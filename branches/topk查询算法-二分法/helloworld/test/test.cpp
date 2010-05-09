// test.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream> 
#include <string>
#include <sstream>
using namespace std;
class kreorderlist
{
public:
	kreorderlist()
	{
		m_ngroupcount = 0;
	}
	~kreorderlist(){}
	int readgroupcount(){return m_ngroupcount;}
	int readitemcount(){return m_nitemcount;}
	int* readgroup(int);
	int putgroup(int*);
	int removegroup(int,int);
	int addgroup(int,int);
	bool checkgroup(int);

private:
	static const char itemsplit = ',';
	static const char groupsplit = ' ';
	int m_ngroupcount;//total group
	int m_nitemcount;//total item in a group
	int m_ngroupserial;//group serial in strlist
	int m_nremovegroupnum;//remove item num
	string m_groupnumlist;//group numnber
	string m_strList;//data serial
	string m_strremovegroup;//store the remove data
};
int* kreorderlist::readgroup(int ngroupnum)
{
	int* pnlist;
	int nlength,groupserial;
	char c;
	string str;
	istringstream stream(m_strList);
	//check whether the group num in the forest
	if(!checkgroup(ngroupnum))return NULL;	
	groupserial = m_ngroupserial;
	while(groupserial-->=0)stream>>str;
	stream.clear();
	stream.str(str);
	stream >> nlength >> c;
	m_nitemcount = nlength;
	pnlist = new int[nlength];
	for (int i = 0; i < nlength; i++)
	{
		stream>>pnlist[i]>>c;
	}
	return pnlist;
}

int kreorderlist::putgroup( int* itemlist)
{
	string s;
	stringstream buf;
	//put the header
	buf<<*itemlist++;
	buf>>s;
	buf.clear();
	m_groupnumlist += s + itemsplit;	
	//put the item
	while(*itemlist != -1)
	{
		buf<<*itemlist++;
		buf>>s;
		buf.clear();
		m_strList += s + itemsplit;		
	}
	m_strList+=groupsplit;
	m_ngroupcount++;
	cout<<m_strList<<endl;
	cout<<m_groupnumlist<<endl;
	cout<<m_ngroupcount<<endl;

	return 1;
}

bool kreorderlist::checkgroup(int ngroupnum)
{
	string str;
	char c;
	istringstream stream(m_groupnumlist);
	int* ngroupnumlist = new int[m_ngroupcount];
	int j = 0;
	while(!stream.eof())
	{
		stream>>ngroupnumlist[j]>>c;
		if (ngroupnumlist[j] == ngroupnum)
		{
			m_ngroupserial = j;
			cout<<m_ngroupserial<<endl;
			return true;
		}
		j++;
	}
	return false;
}

int kreorderlist::removegroup(int ngroupnum, int nserial)
{		
	if (checkgroup(ngroupnum))
	{
		int* pnlist;
		int* pngrouplist;
		int nlength,groupserial;
		char c;
		string str,groupitem;
		stringstream buf;
		istringstream stream(m_strList);
		m_nremovegroupnum = ngroupnum;
		groupserial = m_ngroupserial;
		m_strList = "";//重组m_strList
		for (int i = 0; i < m_ngroupcount; i++)
		{
			stream>>groupitem;
			if (m_ngroupserial != i)
			{
				m_strList += groupitem + groupsplit;
			}
			else
				str = groupitem;
		}
		stream.clear();
		stream.str(str);
		stream >> nlength >> c;
		m_nitemcount = nlength;
		pnlist = new int[nlength + 2];
		//重组m_strremovegroup
		for (int i = 1; i < nlength + 1; i++)
		{
			stream>>pnlist[i]>>c;
		}
		pnlist[0] = ++m_nitemcount;
		pnlist[nlength + 1] = nserial;
		for (int i = 0; i < nlength + 2; i++)
		{
			buf<<pnlist[i];
			buf>>str;
			buf.clear();
			m_strremovegroup += str + itemsplit;	
		}

		//重组m_groupnumlist
		stream.str(m_groupnumlist);
		pngrouplist = new int[m_ngroupcount];
		m_groupnumlist.clear();		
		for (int i = 0; i < m_ngroupcount; i++)
		{
			stream >>pngrouplist[i]>>c;		
			if (pngrouplist[i] != m_nremovegroupnum)
			{
				buf<<pngrouplist[i];
				buf>>str;
				buf.clear();
				m_groupnumlist += str + itemsplit;
			}			
		}
		m_ngroupcount--;//组数减少

	}
	cout << endl <<  m_ngroupcount//total group
		<<" "<< m_nitemcount//total item in a group
		<<" "<< m_ngroupserial//group serial in strlist
		<<" "<< m_nremovegroupnum//remove item num
		<<" "<< m_groupnumlist//group numnber
		<<" "<< m_strList//data serial
		<<" "<< m_strremovegroup;//store the remove data
	return 1;
}

int kreorderlist::addgroup(int ngroupnum, int nserial)
{
	if (!checkgroup(ngroupnum))
	{
		int* pnlist;
		int* pngrouplist;
		int nlength,groupserial;
		char c;
		string str,groupitem;
		stringstream buf;
		istringstream stream(m_strList);
		if (ngroupnum != m_nremovegroupnum)
		{
			//m_groupnumlist重组
			buf<<ngroupnum;
			buf>>str;
			m_groupnumlist += str + itemsplit;
			//m_strList重组
			str = "1";
			m_strList += groupsplit + str;
			buf.clear();
			buf<<nserial;
			buf>>str;
			buf.clear();
			m_strList += itemsplit + str;
			m_ngroupcount++;
		}
		else
		{
			//m_strremovegroup重组
			stream.str(m_strremovegroup);
			stream >> nlength >> c;
		}


	}
	cout << endl <<  m_ngroupcount//total group
		<<" "<< m_nitemcount//total item in a group
		<<" "<< m_ngroupserial//group serial in strlist
		<<" "<< m_nremovegroupnum//remove item num
		<<" "<< m_groupnumlist//group numnber
		<<" "<< m_strList//data serial
		<<" "<< m_strremovegroup;//store the remove data
	return 1;
}
int main() 
{
	kreorderlist a;
	int b[] = {2,4,2,4,34,15,-1};// group serial ,total num, data
	int b1[] = {1,2,1,4,-1};
	int b2[] = {0,2,4,2,-1};
	int b3[] = {3,2,2,4,-1};
	int c[1] ={0};
	int* d;
	a.putgroup(b);
	a.putgroup(b1);
	a.putgroup(b2);
	a.putgroup(b3);
	d = a.readgroup(2);
	if (d == NULL)
	{
		cout << "no";
	}

	int e = a.readitemcount();
	for(int i = 0; i < e; i++)
	{
		cout<<d[i]<<" ";
	}

	a.removegroup(0, 340);
	a.addgroup(0, 341);
	system("PAUSE");
	return 0;
} 

