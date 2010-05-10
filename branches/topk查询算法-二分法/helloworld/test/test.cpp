// test.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <iostream> 
#include <string>
#include <sstream>
using namespace std;
typedef int Status;
class kreorderlist
{
public:
	/* States */
	enum States {Removed, Empty, Added}; // states
	kreorderlist()
	{
		m_ngroupcount = 0;
		m_state = Empty;
	}
	~kreorderlist(){}
	int readgroupcount(){return m_ngroupcount;}
	int readitemcount(){return m_nitemcount;}
	int* readgroup(int);
	int putgroup(int*);
	int enterRemove(int,int);
	int enterAdd(int,int);
	int enterJoin(int, int);
	bool checkgroup(int);
	void putitem(int,int);
	void changeState(States);
	States currentState() { return m_state; }

	/*Events*/
	enum Event {Remove, Join, Add};

private:
	static const char itemsplit = ',';
	static const char groupsplit = ' ';
	States m_state;
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

int kreorderlist::enterRemove(int ngroupnum, int nserial)
{		
	int* pngrouplist;
	int nlength = m_ngroupcount,
		ngroupserial = 0, 
		ncount = 0, 
		nrecord = -1;
	char c;
	string str,groupitem;
	stringstream buf;
	istringstream stream(m_groupnumlist);
	m_nremovegroupnum = ngroupnum;
	m_ngroupcount--;//组数减少
	if (m_strremovegroup != "")
	{
		m_ngroupcount++;//同组元素，+1
		stream.str(m_strremovegroup);
		m_strremovegroup = "";
		stream>>nlength>>c;
		pngrouplist = new int[nlength];
		buf<<(nlength + 1);
		buf>>groupitem;
		groupitem += itemsplit;
		buf.clear();
		for (int i = 0; i < nlength; i++)
		{
			stream>>pngrouplist[i]>>c;
			buf<<pngrouplist[i];
			buf>>str;
			buf.clear();
			groupitem += str + itemsplit;
		}
		buf<<nserial;
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
		m_strremovegroup = groupitem;
		cout << endl <<  m_ngroupcount//total group
			<<"|"<< m_nitemcount//total item in a group
			<<"|"<< m_ngroupserial//group serial in strlist
			<<"|"<< m_nremovegroupnum//remove item num
			<<"|"<< m_groupnumlist//group numnber
			<<"|"<< m_strList//data serial
			<<"|"<< m_strremovegroup;//store the remove data
		return 1;
	}
	//重组m_groupnumlist
	m_groupnumlist = "";
	for (int i = 0; i < nlength; i++)
	{
		stream>>ngroupserial>>c;
		if (ngroupserial != ngroupnum)
		{
			buf<<ngroupserial;
			buf>>str;
			buf.clear();
			m_groupnumlist += str + itemsplit;
		}	
		else
			nrecord = i;
	}
	//重组m_strList
	stream.str(m_strList);
	m_strList = "";
	for (int i = 0; i < nlength; i++)
	{
		stream>>groupitem;
		if (nrecord != i)
		{
			m_strList += groupitem + groupsplit;
		}
		else
			str = groupitem;
	}
	//重组m_strremovegroup
	stream.str(str);
	stream>>nlength>>c;
	pngrouplist = new int[nlength];
	buf<<(nlength + 1);
	buf>>groupitem;
	groupitem += itemsplit;
	buf.clear();
	for (int i = 0; i < nlength; i++)
	{
		stream>>pngrouplist[i]>>c;
		buf<<pngrouplist[i];
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
	}
	buf<<nserial;
	buf>>str;
	buf.clear();
	groupitem += str + itemsplit;
	m_strremovegroup = groupitem;
	
	cout << endl <<  m_ngroupcount//total group
		<<"|"<< m_nitemcount//total item in a group
		<<"|"<< m_ngroupserial//group serial in strlist
		<<"|"<< m_nremovegroupnum//remove item num
		<<"|"<< m_groupnumlist//group numnber
		<<"|"<< m_strList//data serial
		<<"|"<< m_strremovegroup;//store the remove data
	return 1;
}

int kreorderlist::enterAdd(int ngroupnum, int nserial)
{

	int* pngrouplist;
	int nlength = m_ngroupcount,groupserial = -1;	
	char c;
	string str,groupitem;
	stringstream buf;
	istringstream stream(m_groupnumlist);
	m_groupnumlist = "";
	m_ngroupcount++;//组数+1
	pngrouplist = new int[nlength];
	//m_groupnumlist重组
	for (int i = 0; i < nlength; i++)
	{
		stream >>pngrouplist[i]>>c;	
		if (pngrouplist[i] != ngroupnum)
		{
			buf<<pngrouplist[i];
			buf>>str;
			buf.clear();
			m_groupnumlist += str + itemsplit;
		}
		else
		{
			groupserial = i;
			m_ngroupcount--;//是相同的组，组数-1
		}
	}
	buf<<ngroupnum;
	buf>>str;
	buf.clear();
	m_groupnumlist += str + itemsplit;
	//m_strList重组
	stream.clear();
	stream.str(m_strList);
	m_strList = "";
	str="";
	for (int i = 0; i < nlength; i++)
	{
		stream>>groupitem;
		if (groupserial != i)
		{
			m_strList += groupitem + groupsplit;
		}
		else
		{
			str = groupitem;
		}			
	}
	if (str == "")
	{
		nlength = 1;
		groupitem = "";
		buf<<nlength;
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
		buf<<nserial;
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
		m_strList += groupitem + groupsplit;
	}
	else
	{
		stream.str(str);
		str = "";
		stream>>nlength>>c;
		buf<<(nlength + 1);
		buf>>groupitem;
		groupitem += itemsplit;
		buf.clear();
		for (int i = 0; i < nlength; i++)
		{
			stream>>pngrouplist[i]>>c;
			buf<<pngrouplist[i];
			buf>>str;
			buf.clear();
			groupitem += str + itemsplit;
		}
		buf<<nserial;
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
		m_strList += groupitem + groupsplit;
	}

	cout << endl <<  m_ngroupcount//total group
		<<"|"<< m_nitemcount//total item in a group
		<<"|"<< m_ngroupserial//group serial in strlist
		<<"|"<< m_nremovegroupnum//remove item num
		<<"|"<< m_groupnumlist//group numnber
		<<"|"<< m_strList//data serial
		<<"|"<< m_strremovegroup;//store the remove data
	return 1;
}


int kreorderlist::enterJoin(int ngroupnum,int nserial)
{
	string str;
	stringstream buf;
	m_ngroupcount++;
	m_strList += m_strremovegroup + groupsplit;
	buf<<m_nremovegroupnum;
	buf>>str;
	buf.clear();
	m_groupnumlist += str;
	m_strremovegroup = "";
	m_nremovegroupnum = -1;
	cout << endl <<  m_ngroupcount//total group
		<<"|"<< m_nitemcount//total item in a group
		<<"|"<< m_ngroupserial//group serial in strlist
		<<"|"<< m_nremovegroupnum//remove item num
		<<"|"<< m_groupnumlist//group numnber
		<<"|"<< m_strList//data serial
		<<"|"<< m_strremovegroup;//store the remove data

	return 1;
}


void kreorderlist::putitem(int ngroupnum,int nserial)
{
	switch (m_state)
	{
	case Empty:
		enterAdd(ngroupnum, nserial);
		changeState(Added);
		break;
	case Added:
		if (checkgroup(ngroupnum))
		{
			enterRemove(ngroupnum, nserial);
			changeState(Removed);
		}
		else
			enterAdd(ngroupnum, nserial);
		break;
	case Removed:
		if (ngroupnum == m_nremovegroupnum)
		{
			enterRemove(ngroupnum, nserial);
		}
		else
		{
			enterJoin(ngroupnum, nserial);
			enterAdd(ngroupnum, nserial);
			changeState(Added);
		}
		break;
	}

}

void kreorderlist::changeState( States state)
{
	m_state = state;
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
	// 	a.putgroup(b);
	// 	a.putgroup(b1);
	// 	a.putgroup(b2);
	// 	a.putgroup(b3);
	

	// 	d = a.readgroup(2);
	// 	if (d == NULL)
	// 	{
	// 		cout << "no";
	// 	}

	// 	int e = a.readitemcount();
	// 	for(int i = 0; i < e; i++)
	// 	{
	// 		cout<<d[i]<<" ";
	// 	}
	a.putitem(4, 341);
	a.putitem(4, 341);
	a.putitem(1, 342);
	a.putitem(1, 342);
	a.putitem(1, 342);	
	a.putitem(2, 342);
	a.putitem(3, 342);
	a.putitem(3, 342);
	a.putitem(1, 342);
	a.putitem(2, 342);
	a.putitem(2, 342);
	a.putitem(3, 342);
	a.putitem(3, 342);
	a.putitem(4, 341);
	
	system("PAUSE");
	return 0;
} 

