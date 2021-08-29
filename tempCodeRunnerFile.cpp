#include<iostream>
using namespace std;

void count(int char[],int b){
    cout<<char[0];
    // int size = sizeof(char)/sizeof(char[0]);
    cout<<size;
}

int main(){
    int n = 5;
    int arr = {1,2,3,4,5,6};

    count(arr[],6);
}