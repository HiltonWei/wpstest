// helloworld.cpp : Defines the entry point for the console application.
//
#include <iostream>
#include <cstdlib>
#include <time.h>
#include <math.h>
#define MAXN 10000
#define LIMIT 500
using namespace std;
int a[MAXN];
int last;
int count_sw, count_cmp;

//初始化测试数据
void init () {
	srand(time(0));
	//参与查询的数据
	for(int i=0; i<MAXN; ++i) {
		a[i] =  rand()%MAXN;
	}
	return ;
}
void swap(int &x, int &y) {
	int t ;
	t = x;
	x = y;
	y = t;
	return ;
}

void qsort(int arr[], int start, int end) {
	if(start < end) {
		int mid = arr[rand()%(end-start) + start];
		int i = start - 1;
		int j = end + 1;
		while (true) {
			while (arr[++i] > mid) count_cmp++;
			while (arr[--j] < mid) count_cmp++;
			if(i>=j) break;
			swap(arr[i], arr[j]);
			count_sw++;
		}
		qsort (arr, start, i-1);
		qsort (arr, j+1, end);
	}
	return ;
}

void topK(int arr[], int start, int end, int k) {
	if(start < end) {
		int mid = arr[rand()%(end-start) + start];
		int i = start - 1;
		int j = end + 1;
		while (true) {
			while (arr[++i] > mid) count_cmp++;
			while (arr[--j] < mid) count_cmp++;
			if(i>=j) break;
			swap(arr[i], arr[j]);
			count_sw++;
		}
		if(i-start > k) {
			last = i-1;
			topK (arr, start, i-1, k);
		} else{
			qsort(arr, start, last);
		}
	}
	return ;
}

int main() {
	init();
	count_sw = 0;
	count_cmp = 0;
	topK(a, 0, MAXN-1,10);
	cout << "Result:" << MAXN << endl;
	for(int i=0; i<10; ++i) {
		cout << a[i] << endl;
	}
	cout << "Swaps:" << count_sw << endl;
	cout << "Campare:" << count_cmp << endl;
	int n;
	cin>>n;
	return 0;
}

