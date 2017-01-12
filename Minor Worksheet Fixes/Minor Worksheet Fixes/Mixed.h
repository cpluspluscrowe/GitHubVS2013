#include <queue>
void writeOnFinish(std::string finishedFile);
void writeError(std::string errorFile);
void fillFileQueue(std::queue<std::string>* files,std::string path);
