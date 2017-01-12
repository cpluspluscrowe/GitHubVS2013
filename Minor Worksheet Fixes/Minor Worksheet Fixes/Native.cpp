#include "Native.h"
#include "Mixed.h"
#include <queue>
#include <string>
#include <fstream>
#include <iostream>
using namespace std;

//This adds filepaths to the files queue
void fillFileQueue(std::queue<std::string>* files,std::string path){
	if (fs::is_directory(path)){
		for (fs::recursive_directory_iterator it(path), eit; it != eit; ++it){
			if (!fs::is_directory(it->path())){
				files->push(it->path().string());
			}
		}
	}
}

void writeError(std::string errorFile){
	ofstream myfile;
	myfile.open("C:\\Users\\CCrowe\\Desktop\\errors.txt", ios::app);
	myfile << errorFile << "\n";
	myfile.close();
}
void writeOnFinish(std::string finishedFile){
	ofstream myfile;
	myfile.open("C:\\Users\\CCrowe\\Desktop\\finished.txt", ios::app);
	myfile << finishedFile << "\n";
	myfile.close();
}