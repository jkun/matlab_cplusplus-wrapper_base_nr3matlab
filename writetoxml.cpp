#include"nr3matlab.h"
#include"libxl.h"
using namespace libxl;
void help();
void writedata(string filename,MatDoub &indata);
void mexFunction(int nlhs,mxArray *plhs[],int nrhs,const mxArray *prhs[]){
	string e=MatString(prhs[0]);
	string filename;
	if(e=="-h" || e=="help"){
		help();
		return;
	}
	cout<<"begin to insert into excel file\n";
	MatDoub indata(prhs[1]);
	writedata(e,indata);
}
void help(){
	cout<<"Matlab data to excel program"<<endl;
	cout<<"parameters as follow:"<<endl;
	cout<<"-h or help: help list";
	cout<<"the command line will be two class:"<<endl;
	cout<<"-h or ";
	cout<<"filename data "<<endl;
}
void writedata(string filename,MatDoub &indata){
	Book* book = xlCreateBook(); // use xlCreateXMLBook() for working with xlsx files
	Sheet* sheet = book->addSheet("Sheet1");
	int m=indata.nrows(),n=indata.ncols();
	for(int i=0;i<m;i++)
		for(int j=0;j<n;j++)
			sheet->writeNum(j+2, i+2, indata[i][j]);
	book->save(filename.c_str());
	book->release();
}