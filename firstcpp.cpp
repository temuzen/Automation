#include<iomanip>
#include<iostream>
#include<fstream>
#include<string>

using namespace std;
int main()
{
    ofstream outData;
    string s;
    cin>>s;

    int n = 1;
    string x = "=RTD(pi.rtdserver,,""MCX_ALUMINI18SEPFUT,""TradingSymbol)";
    outData.open("outfile.csv",ios::app);

    outData<<x<<endl;
    outData<<n<<endl;
    system("pause");
    return 0;

    





}
