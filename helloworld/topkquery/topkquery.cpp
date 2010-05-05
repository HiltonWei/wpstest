// helloworld.cpp : Defines the entry point for the console application.
//
#include <iostream>
#include <cstdlib>
#include <time.h>
#include <math.h>
#include "hcode.h"
#include "kundata.h"
#include "ktopquery.h"
using namespace std;

//初始化测试数据

int main() {
	cout << *(ktopquery::GetSingleton());
	
	cout << *(koption::GetSingleton());

	int n;
	cin >> n;
	return 0;
}

