# -*- coding: utf-8 -*-
"""
Created on Thu Dec  8 20:25:14 2022

@author: Jose
"""

from ortools.linear_solver import pywraplp
from IOfunctionsExcel import *




name = "ORoom Data.xlsx"
excel_document = openpyxl.load_workbook(name, data_only=True)
sheetPat = excel_document['Patients']
sheetSurg = excel_document['Surgeons']
sheetOR = excel_document['ORs']
sheetTime = excel_document['Time_Slots']
sheetDays = excel_document['Days']
sheetCalendario = excel_document['Calendario']
sheetCirujanos = excel_document['Cirujanos']

# Se lee la tabla de quirofanos con sus posibles operaciones

D_Quirofanos = Read_Excel_to_NesteDic(sheetOR, ['A1', 'M5'][0], ['A1', 'M5'][1])

# Se lee la tabla de cirujanos con sus posibles operaciones

D_Cirujanos = Read_Excel_to_NesteDic(sheetSurg, ['B4', 'N10'][0], ['B4', 'N10'][1])

# Se lee la tabla de pacientes con su importancia, dia de admisión y tipo de cirugía

D_Pacientes = Read_Excel_to_NesteDic(sheetPat, ['A1', 'D101'][0], ['A1', 'Q101'][1])



L_Cirujanos = [c for c in range(0,6)]

#L_Pacientes = [p for p in D_Pacientes]
L_Pacientes = [p for p in range(0,100)]

L_Pacientes2 = []
for p in L_Pacientes:
    L_Pacientes2.append('#p#%d'%(p+1))
    
L_Retrasos = []
for p in L_Pacientes2:
    L_Retrasos.append(D_Pacientes[p]['retraso'])
    
L_Gravedad = []
for p in L_Pacientes2:
    L_Gravedad.append(D_Pacientes[p]['imp']) 

    
L_Factor = []
for p in L_Pacientes:
    if L_Gravedad[p] == 1:
        L_Factor.append(1)
    if L_Gravedad[p] == 2:
        L_Factor.append(500)
    if L_Gravedad[p] == 3:
        L_Factor.append(3000)
    if L_Gravedad[p] == 4:
        L_Factor.append(10000)
    
      
L_Quirofanos = [q for q in range(0,4)]

L_Slots = [t for t in range(0,15)]

L_Dias = [d for d in range(0,5)]

L_Tipo_Cirugia = ['sk1', 'sk2', 'sk3', 'sk4', 'sk5', 'sk6', 'sk7', 'sk8', 'sk9', 'sk10', 'sk11', 'sk12']

#ToStr={o:i for o,i in enumerate(L_Tipo_Cirugia)}  # Diccionario que convierte número de tipo de operación a su label
#ToInt={i:o for o,i in enumerate(L_Tipo_Cirugia)}  # Diccionario que convierte Label del tipo de operación a su número


def Problema():

    
    modelname = 'Hospital'
    solver = pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING
    
    solver = pywraplp.Solver(modelname, solver)
    
    x = [[[[solver.IntVar(0, 1, 'X_%s_%s_%s_%s' % (p, c, q, t)) for t in L_Slots]  for q in L_Quirofanos] for c in L_Cirujanos] for p in L_Pacientes] #Paciente p, que es operado por cirujano c, que se opera en el Quirofano q y en el Slot t           
    #H = [solver.NumVar(0, 15 * 2, 'H_%s' %(h, d)) for i in L_Cirujanos]
    deltac = [[solver.IntVar(0,1, 'deltac_%s_%s' %(p, c)) for c in L_Cirujanos] for p in L_Pacientes] 
    deltaq = [[solver.IntVar(0,1, 'deltaq_%s_%s' %(p, q)) for q in L_Quirofanos] for p in L_Pacientes] 
    
    print('Número de variables = ',solver.NumVariables())
    
    #Restricción para que un cirujano, en un quirófano y en un horario no atienda a más de un paciente

    for c in L_Cirujanos:
        for q in L_Quirofanos:        
            for t in L_Slots:
                solver.Add(sum(x[p][c][q][t] for p in L_Pacientes) <= 1)
        
    print('Nueva restricción')
    
    #Restricción para que un paciente, en un quirófano y en un horario no es atendido por más de un cirujano
    
    for p in L_Pacientes:  
        for q in L_Quirofanos: 
            for t in L_Slots:
                solver.Add(sum(x[p][c][q][t] for c in L_Cirujanos) <= 1)
        
    print('Nueva restricción')
    
    #Restricción para que un paciente, atendido por un cirujano y en un horario no es operado en varios quirofanos  
    
    for p in L_Pacientes:
        for q in L_Quirofanos:
            for c in L_Cirujanos:
                solver.Add(sum(x[p][c][q][t] for t in L_Slots) <= 1)
        
    print('Nueva restricción')
    
    #Restricción para que un paciente, atendido por un cirujano y en un quirofano no es operado en varios horarios
    
    for p in L_Pacientes:
        for t in L_Slots:
            for c in L_Cirujanos:
                solver.Add(sum(x[p][c][q][t] for q in L_Quirofanos) <= 1)
        
    print('Nueva restricción')
    
    #Restricción para que un paciente p solo se opere por un cirujano c, un quirofano q y en un slot t
    for p in L_Pacientes:
        solver.Add(sum(x[p][c][q][t] for c in L_Cirujanos for q in L_Quirofanos for t in L_Slots) <= 1)
            
    print('Nueva restricción')
    
    
    #Restricción para que solo se opere en un slot t en un quirofano q
    for t in L_Slots:
        for q in L_Quirofanos:
            solver.Add(sum(x[p][c][q][t] for p in L_Pacientes for c in L_Cirujanos) <= 1)
            
    print('Nueva restricción')
    
    #Restricción para que solo se opere en un slot t un cirujano c
    for t in L_Slots:
        for c in L_Cirujanos:
            solver.Add(sum(x[p][c][q][t] for p in L_Pacientes for q in L_Quirofanos) <= 1)
            
    print('Nueva restricción')
    

    print('Número de restricciones = ', solver.NumConstraints())
        
   
    #Restricciones para definir variable deltac: 1 si paciente p puede ser operado por cirujano c según skills
    for c in L_Cirujanos:    
        for p2 in L_Pacientes2:
            for s in L_Tipo_Cirugia:
                #print(L_Pacientes2.index(p2))
                if D_Pacientes[p2][s] == 1:
                    if D_Cirujanos[c][s] == 1:
                        solver.Add(deltac[L_Pacientes2.index(p2)][c] == 1)
                        #print('Cirujano', c ,' sabe operar a este paciente', p2)
                    else:
                        solver.Add(deltac[L_Pacientes2.index(p2)][c] == 0)
                        #print('Este cirujano', c,'no puede operar a este paciente', p2)

    #Restricciones para definir variable deltaq: 1 si paciente p puede ser operado en quirofano q según skills
    for q in L_Quirofanos:    
        for p2 in L_Pacientes2:
            for s in L_Tipo_Cirugia:
                if D_Pacientes[p2][s] == 1:
                    if D_Quirofanos[q][s] == 1:
                        solver.Add(deltaq[L_Pacientes2.index(p2)][q] == 1)
                        #print('En quirofano', q,' se puede operar a este paciente', p2)
                    else:
                        solver.Add(deltaq[L_Pacientes2.index(p2)][q] == 0)
                        #print('En quirofano', q,'no se puede operar a este paciente', p2)

    #Restricciones para determinar que cirujano c puede operar a paciente p 
    for p in L_Pacientes:
        for c in L_Cirujanos:
            solver.Add(sum(x[p][c][q][t] for q in L_Quirofanos for t in L_Slots) <= deltac[p][c])

    #Restricción para determinar en que quirofano q se puede operar al paciente p
    for p in L_Pacientes:
        for q in L_Quirofanos:
            solver.Add(sum(x[p][c][q][t] for c in L_Cirujanos for t in L_Slots) <= deltaq[p][q])


    print('Número de restricciones = ', solver.NumConstraints())
    
    
    #Restricción para maximo de horas operando por cirujano c
    for c in L_Cirujanos:
        solver.Add((2 * sum(x[p][c][q][t] for q in L_Quirofanos for t in L_Slots for p in L_Pacientes)) <= 18)
        
    print('Nueva restricción')
    
    #Restricción para min de horas operando por cirujano c
    for c in L_Cirujanos:
        solver.Add((2 * sum(x[p][c][q][t] for q in L_Quirofanos for t in L_Slots for p in L_Pacientes)) >= 10)
        
    print('Nueva restricción')
    
    
    #Restricción para que se opere al menos al 45% de la lista de pacientes
    solver.Add(sum(x[p][c][q][t] for q in L_Quirofanos for c in L_Cirujanos for t in L_Slots for p in L_Pacientes) >= len(L_Pacientes) * 0.45)

    
    print('Número de restricciones = ', solver.NumConstraints())     
    
    
    
    solver.Maximize(solver.Sum(L_Retrasos[p] * L_Factor[p] * x[p][c][q][t] for q in L_Quirofanos for c in L_Cirujanos for t in L_Slots for p in L_Pacientes))
    
    
    status = solver.Solve() 
    if status==pywraplp.Solver.OPTIMAL:
        print('El problema tiene solucion')
    
    print('\n')
    
    
    cantPacientesOperados = 0
    #Imprimir soluciones por pantalla
    for p2 in L_Pacientes2:
        for c in L_Cirujanos:
            for q in L_Quirofanos:
                for t in L_Slots:
                    if x[L_Pacientes2.index(p2)][c][q][t].solution_value() == 1:
                        cantPacientesOperados += 1
                        print('Paciente = %s, Cirujano = %d, Quirófano = %d, Slot = %d'%(p2, c, q, t), [x[L_Pacientes2.index(p2)][c][q][t].solution_value()])
    
    
    print('\n')
    print('Cantidad de pacientes operados =', cantPacientesOperados)
    print('\n')
    print('Porcentaje de pacientes operados de la lista de espera =', (cantPacientesOperados/len(L_Pacientes)) * 100, '%')
    print('\n')
    
    horasPorCir = []
    
    #Imprimir soluciones por pantalla
    for c in L_Cirujanos:
        horasPorCir = [2 * sum(x[p][c][q][t].solution_value() for q in L_Quirofanos for t in L_Slots for p in L_Pacientes)]
        Write_List_to_Excel(excel_document, name, sheetCirujanos, horasPorCir, 'a'+str(c+2),'a'+str(c+2))
        Write_List_to_Excel(excel_document, name, sheetCirujanos, L_Cirujanos, 'b2','b7')
        print('Horas en las que ha operado el cirujano %d'%( c), [horasPorCir])
    
    #print(horasPorCir)
    
    #Write_List_to_Excel(excel_document, name, sheetCirujanos, horasPorCir, 'a2','a7')
   
 
    """
    Pacientes = []
    Quirofanos = []
    Slots = []
    Cirujanos = []
    
    #Escribir soluciones en Excel, se ha variado el rango de slots en diferentes ejecuciones de codigo para escribir en el Excel
    
    for c in L_Cirujanos:
            for p2 in L_Pacientes2:
                for q in L_Quirofanos:
                    for t in range(0,15):
                        if x[L_Pacientes2.index(p2)][c][q][t].solution_value() == 1:
                            Pacientes.append(p2)
                            Cirujanos.append(c)
                            Quirofanos.append(q)
                            Slots.append(t)
    
    
    Write_List_to_Excel(excel_document, name, sheetCirujanos, Cirujanos, 'd2', 'd55')
    Write_List_to_Excel(excel_document, name, sheetCirujanos, Quirofanos, 'e2', 'e55')
    Write_List_to_Excel(excel_document, name, sheetCirujanos, Slots, 'f2', 'f55')
    Write_List_to_Excel(excel_document, name, sheetCirujanos, Pacientes, 'g2', 'g55')
    """
Problema()






