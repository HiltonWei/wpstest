// helloworld.cpp : Defines the entry point for the console application.
//
#include <iostream>
#include <cstdlib>
#include <time.h>
#include <math.h>
#include "hcode.h"
#include "kundata.h"
#include "ktopquery.h"
#include <Windows.h>
using namespace std;

//初始化测试数据

int main() {
	LARGE_INTEGER m_nFreq;
	LARGE_INTEGER m_nBeginTime;
	LARGE_INTEGER nEndTime;

	QueryPerformanceFrequency(&m_nFreq); // 获取时钟周期
	QueryPerformanceCounter(&m_nBeginTime); // 获取时钟计数

	cout << *(ktopquery::GetSingleton());
	QueryPerformanceCounter(&nEndTime);
	cout << "Run tiem : "
		<<(nEndTime.QuadPart-m_nBeginTime.QuadPart)*1000.0/m_nFreq.QuadPart
		<< " ms"<<endl;

	
	cout << *(koption::GetSingleton());

	int n;
	cin >> n;
	return 0;
}

