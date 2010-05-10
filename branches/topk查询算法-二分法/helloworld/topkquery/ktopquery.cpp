#include "ktopquery.h"
#include <iostream> 
#include <string>
#include <sstream>
using namespace std;
ktopquery* ktopquery::instance = NULL;
kreorderlist* kreorderlist::instance = NULL;
ktopquery::ktopquery(void)
{
	init();
}


ktopquery::~ktopquery(void)
{
}

int ktopquery::init()
{
	count_sw = 0;
	count_cmp = 0;
	m_pkdata = kundata::GetSingleton();
	m_pdatalist = m_pkdata->readdatalist();
	m_pprolist = m_pkdata->readprobabilitylist();
	m_pkitemlist = m_pkdata->readitemlist();
	arr = new int[koption::s_nMaxnum];
	for (int i = 0; i < koption::s_nMaxnum; i++)
	{
		arr[i] = i;
	}
	topK(0, koption::s_nMaxnum-1);
	for (int i = 0; i < koption::s_nMaxnum; i++)
	{
		compression(i);
		subsetprovalue(i);
	}
	
	return 1;
}

void ktopquery::swap( int& x,int& y)
{
	int t ;
	t = x;
	x = y;
	y = t;
	return ;
}

void ktopquery::qsort(int start, int end )
{
	if(start < end) {
		int mid = arr[rand()%(end-start) + start];
		int i = start - 1;
		int j = end + 1;
		while (true)
		{
			while ((m_pkitemlist[arr[++i]]).readdata() > (m_pkitemlist[mid]).readdata()) count_cmp++;
			while ((m_pkitemlist[arr[--j]]).readdata() < (m_pkitemlist[mid]).readdata()) count_cmp++;
			if(i>=j) break;
			swap(arr[i], arr[j]);
			count_sw++;
		}
		qsort (start, i-1);
		qsort (j+1, end);
	}
	return ;
}

void ktopquery::topK(int start, int end)
{
	int k = koption::s_nK;
	if(start < end) {
		int mid = arr[rand()%(end-start) + start];
		int i = start - 1;
		int j = end + 1;
		while (true) {
			while (m_pkitemlist[arr[++i]].readdata() > m_pkitemlist[mid].readdata()) count_cmp++;
			while (m_pkitemlist[arr[--j]].readdata() < m_pkitemlist[mid].readdata()) count_cmp++;
			if(i>=j) break;
			swap(arr[i], arr[j]);
			count_sw++;
		}
		if(i-start > k) {
			last = i-1;
			topK (start, i-1);
		} else{
			qsort(start, last);
		}
	}
	return ;
}

int ktopquery::compression(int i)
{
	if (i = 1)return i;

	if (!m_pkitemlist[arr[i]].readmark())return -1;

}

void ktopquery::subsetprovalue(int i)
{

}
std::ostream & operator <<(std::ostream &os,ktopquery &item)
{
	for(int i = 0; i < koption::s_nK; i++) 
	{
		os << item.m_pkitemlist[item.arr[i]];
	}
	os << "Swaps:" << item.count_sw << endl;
	os << "Campare:" << item.count_cmp << endl;
	return os;
}

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
			return true;
		}
		j++;
	}
	return false;
}
std::ostream & operator <<(std::ostream &out,kreorderlist &item)
{
	cout << endl <<  item.m_ngroupcount//total group
		<<"|"<< item.m_ngroupserial//group serial in strlist
		<<"|"<< item.m_nremovegroupnum//remove item num
		<<"|"<< item.m_groupnumlist//group numnber
		<<"|"<< item.m_strList//data serial
		<<"|"<< item.m_strremovegroup;//store the remove data
	return out;
}

int kreorderlist::enterRemove(int ngroupnum, int nserial)
{		
	int nlength = m_ngroupcount,
		ngroupserial = 0, 
		nrecord = -1,
		nread = 0;
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
		buf<<(nlength + 1);
		buf>>groupitem;
		groupitem += itemsplit;
		buf.clear();
		for (int i = 0; i < nlength; i++)
		{
			stream>>nread>>c;
			buf<<nread;
			buf>>str;
			buf.clear();
			groupitem += str + itemsplit;
		}
		buf<<nserial;
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
		m_strremovegroup = groupitem;
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
	stream.clear();
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
	stream.clear();
	stream.str(str);
	stream>>nlength>>c;
	buf<<(nlength + 1);
	buf>>groupitem;
	groupitem += itemsplit;
	buf.clear();
	for (int i = 0; i < nlength; i++)
	{
		stream>>ngroupserial>>c;
		buf<<ngroupserial;
		buf>>str;
		buf.clear();
		groupitem += str + itemsplit;
	}
	buf<<nserial;
	buf>>str;
	buf.clear();
	groupitem += str + itemsplit;
	m_strremovegroup = groupitem;

	return 1;
}

int kreorderlist::enterAdd(int ngroupnum, int nserial)
{
	int nlength = m_ngroupcount,
		groupserial = -1,
		nread = 0;	
	char c;
	string str,groupitem;
	stringstream buf;
	istringstream stream(m_groupnumlist);
	m_groupnumlist = "";
	m_ngroupcount++;//组数+1
	//m_groupnumlist重组
	for (int i = 0; i < nlength; i++)
	{
		stream >>nread>>c;	
		if (nread != ngroupnum)
		{
			buf<<nread;
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
			stream>>nread>>c;
			buf<<nread;
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
			if (checkgroup(ngroupnum))
			{
				enterRemove(ngroupnum, nserial);
			}
			else
			{
				enterAdd(ngroupnum, nserial);
				changeState(Added);
			}
		}
		break;
	}

}

void kreorderlist::changeState( States state)
{
	m_state = state;
}