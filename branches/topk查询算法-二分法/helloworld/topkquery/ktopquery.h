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
	int m_ngroupcount;//总数
	int m_nremovegroupnum;//被移除的原则编号
	string m_groupnumlist;//元组编号
	string m_strList;//数据列
	string m_strremovegroup;//存储被移除的组
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
	string m_strRuleFail;//记录同一规则的概率最大值
	int last;
	double m_dMarkindependfailpro;//记录独立元组的失败的概率最大值
	int m_nCountPrint;//记录显示输出多少个元组
	double m_dTotalOutPr;//输出元素的总概率
	int* arr;//记录排序后的序号
	double* reorderarr;//记录重排后的序号
	int count_sw, count_cmp;//记录topk算法的交换和比较
	int* m_pdatalist;//记录得分值
	int* m_pruleseriallist;//记录得分值
	double* m_pprolist;//记录概率值
	double* m_pruletotalprolist;//记录总概率值
	static ktopquery *instance;//实体
	ktopquery(void);
	~ktopquery(void);
	kdataitem* m_pkitemlist;//记录元组列表
	kundata* m_pkdata;//获取数据用的对象指针
};
