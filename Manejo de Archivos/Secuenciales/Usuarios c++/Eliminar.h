#include <cstdlib>
#include <iostream>
#include <fstream>
#include <conio.h>
using namespace std;
void Eliminar()
{
int MenuIni=71,i=8,Cantidad=0,Tecla;;
bool Encontrado=false;
char Contenido[255];
char Nom[255],Ape[255],Dir[255],Tel[255],Buscar[255];
string busc,nomb;

stringstream Convert;
stringstream Convert1;

ofstream   EscribeTemp("Temp.bkup");
ifstream   Lee("Usuarios.txt");/*

ofstream   EscribeTemp("Temp.bkup");
ofstream   EscribeUser("Usuarios.txt");
*/

Text(MenuIni,13,2);cout<<"================================================";
Text(MenuIni,13,3);cout<<"            *BUSQUEDA DE USUARIOS*              ";
Text(MenuIni,13,4);cout<<"================================================";

Text(MenuIni,25,6); cout<<"Nombre(s):";
Text(MenuIni,35,6);cin.getline(Buscar,255);

//Conversion de char a string
Convert<<Buscar;
Convert>>busc;
//////////////////////

//Muestra en Pantalla Los Nombres de Los Campos
Text(MenuIni,1,i);
cout<<setw(20)<<setfill('-')<<left<<"NOMBRE(S)"<<setw(20);
cout<<"APELLIDOS"<<setw(20)<<"DIRECCION"<<setw(15)<<"TEL"<<endl;    


//Leee El Primer Registro del Archivo
  Lee.getline(Nom,255);
   
    
   while(!Lee.eof())   
   {
  
               
                    
     if(Nom==busc){
                            
     Lee.getline(Ape,255);
     Lee.getline(Dir,255);
     Lee.getline(Tel,255);
                              
       Text(MenuIni,1,i+2);
       Encontrado=true;     
       Cantidad++;      
      
      //Mustra en Pantalla los Datos
      cout<<setw(20)<<setfill(' ')<<left<<Nom<<setw(20);
      cout<<Ape<<setw(20)<<Dir<<setw(15)<<Tel<<endl; 
      
      Text(MenuIni,1,i+7);
      cout<<setw(75)<<setfill('-')<<left<<""<<endl;
      ///////////////////////////////
      
      //Aparatir de aquicomienza el procedimiento para eliminar
      Text(MenuIni,1,i+4);
      cout<<"Desea Eliminar [S/N]:"; 
      Text(MenuIni,22,i+4);Tecla=getch();   
      
      if(Tecla==110|Tecla==78)//si presiona "N"
      {
      Text(MenuIni,1,i+4);
      cout<<"                     "; 
      //Escribe el registro en el archivo temporal
      EscribeTemp<<Nom<<endl;
      EscribeTemp<<Ape<<endl;
      EscribeTemp<<Dir<<endl;
      EscribeTemp<<Tel<<endl;
      
      }    
     if(Tecla==83|Tecla==115)// si presiona "S"
     {  
        //Escribe los datos al Archivo temporal                    

      
      
             
     }      
       
      // Aqui termina el procedimiento para eliminar  
       
       
     }else{
          
                          
     //Lee Otro Registro en Caso de Que no 
     //Coinsida Con el Nombre a Buscar
       
      Lee.getline(Ape,255);
      Lee.getline(Dir,255);
      Lee.getline(Tel,255);         
      
      
      EscribeTemp<<Nom<<endl;
      EscribeTemp<<Ape<<endl;
      EscribeTemp<<Dir<<endl;
      EscribeTemp<<Tel<<endl;
      
      
      }
     // Lee el nombre del siguiente registro para contunuar
     //con la comparacion
     Lee.getline(Nom,255); 
     Convert1<<Nom;
     Convert1>>nomb;

  }
Lee.close();
EscribeTemp.close();


Text(MenuIni,1,i+7);
cout<<setw(75)<<setfill('-')<<left<<""<<endl;
  
Text(MenuIni,1,i+9);
if (Encontrado==false)
{
cout<<"Resultado de la Busqueda: No Se Encontro Usuario!!"<<endl<<endl;
}
else
{
cout<<"Resultado de la Busqueda: "<<Cantidad<<endl<<endl;
}  
cout<<"  PRESIONE CUALQUIER TECLA.....";
cin.get();

//Apartir de Aqui Ace la Actualizacion del Archivo
ofstream   EscribeUser("Usuarios.txt");
ifstream   LeeTep("Temp.bkup");

LeeTep.getline(Nom,255);
LeeTep.getline(Ape,255);
LeeTep.getline(Dir,255);
LeeTep.getline(Tel,255);

while(!LeeTep.eof())   
   {
    EscribeUser<<Nom<<endl;
    EscribeUser<<Ape<<endl;
    EscribeUser<<Dir<<endl;
    EscribeUser<<Tel<<endl;
    LeeTep.getline(Nom,255);
    LeeTep.getline(Ape,255);
    LeeTep.getline(Dir,255);
    LeeTep.getline(Tel,255);                  
                       
   }
LeeTep.close();
EscribeUser.close();

/////////////////////////////////////////////////




}
