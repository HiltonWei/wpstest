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
	int readitem(int, int*&, int&);
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
		m_nremovegroupnum = -1;
		m_state = Empty;
		m_strList = "";
		m_groupnumlist = "";
		m_strremovegroup = "";
	}
	~kreorderlist(){}
	static const char itemsplit = ',';
	static const char groupsplit = ' ';
	States m_state;
	int m_ngroupcount;//����
	int m_nremovegroupnum;//���Ƴ���ԭ����
	string m_groupnumlist;//Ԫ����
	string m_strList;//������
	string m_strremovegroup;//�洢���Ƴ�����
};
class kfailrule
{
public:
	static int putitem(int,double);
	static bool checkitem(int,double);
private:
	static const char itemsplit = ',';
	static const char groupsplit = ' ';
	static string s_strlist;
	static int s_nlength;
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
	double Pr(int);
	double Prk(int);
	double PrSt(int,int);
	void swap(int&,int&);
	void qsort(int start, int end);
	void topK(int start, int end);
	int compression(int);
	void subsetprovalue(int);
	bool prune(int);
	friend std::ostream & operator <<(std::ostream &,ktopquery &); 
private:
	string m_strRuleFail;//��¼ͬһ����ĸ������ֵ
	int last;
	double m_dMarkindependfailpro;//��¼����Ԫ���ʧ�ܵĸ������ֵ
	int m_nCountPrint;//��¼��ʾ������ٸ�Ԫ��
	double m_dTotalOutPr;//���Ԫ�ص��ܸ���
	int* arr;//��¼���������
	double* reorderarr;//��¼���ź�����
	int count_sw, count_cmp;//��¼topk�㷨�Ľ����ͱȽ�
	int* m_pdatalist;//��¼�÷�ֵ
	int* m_pruleseriallist;//��¼�÷�ֵ
	double* m_pprolist;//��¼����ֵ
	double* m_pruletotalprolist;//��¼�ܸ���ֵ
	static ktopquery *instance;//ʵ��
	ktopquery(void);
	~ktopquery(void);
	kdataitem* m_pkitemlist;//��¼Ԫ���б�
	kundata* m_pkdata;//��ȡ�����õĶ���ָ��
};
