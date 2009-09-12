Attribute VB_Name = "Variables"
Public Assigned() As Boolean        'Indica si los nodos ya están asignados
Public BestAnt() As Integer          'Mejor hormiga de cada iteración
Public BestCovered() As Double      'Mejor valor de la Funcion Objetivo
Public BestCused() As Integer       'Capacidad usada de cada vehículo (Mejor solución)
Public BestNv As Integer            'Número de vehículos en la mejor solución
Public BestSolution() As Double     'Mejor Solucion
Public Capv As Integer              'Capacidad de los vehículos
Public Chosen As Integer            'Variable que indica el nodo elegido en algún paso del algoritmo
Public Covered() As Double          'Distancia recorrida por cada vehículo
Public CoveredAnt() As Double
Public Ctrl() As Integer            'Variable de control que indica la posición inicial y final de cada vehículo en el vector solución
Public CtrlAnt() As Integer         'Variable de control que indica la posición inicial y final de cada vehículo en el vector solución en la mejor solución de cada iteración
Public Cused() As Integer           'Capacidad usada de cada vehículo
Public CusedAnt() As Integer
Public CusedAcc As Integer          'Capacidad acumulada usada de los vehículos programados
Public DeltaTao() As Double         'Parámetro propio del método
Public Dem() As Integer             'Demanda de cada nodo
Public Demand As Integer            'Demanda total - Suma
Public Dist() As Double             'Distancia entre cada par de nodos
Public Eta() As Double              'Información heurística - parámetro propio del método
Public First() As Double            '??? Ni idea - Si funciona le buscamos la lógica
Public Furthest As Integer          'Nodo más lejano
Public GenerationNext               'Indice de generaciones
Public K As Integer                 'Contador que representa la indexacion de cada hormiga
Public Max As Double                'Máximo
Public MaxNv As Integer             'Número máximo de vehículos
Public MaxX As Integer              'Maxima coordenada en X (para efectos graficos)
Public MaxY As Integer              'Maxima coordenada en Y (para efectos graficos)
Public MinX As Integer              'Minima coordenada en X (para efectos graficos)
Public MinY As Integer              'Minima coordenada en Y (para efectos graficos)
Public nAnts As Integer             'Número de hormigas
Public Nearest() As Integer         'Nodo más cercano (inicio o final de la ruta) a cada nodo
Public nGen As Integer              'Número de generaciones
Public Nodes As Integer             'Número de nodos
Public NProblem As Integer          'Número del problema
Public Nv As Integer                'Número de vehículos
Public Prob As Double               'Variable auxiliar para el cálculo de probabilidades
Public Problem As String            'Problema(s) a resolver
Public Random As Double             'Variable auxiliar para el uso de números aleatorios
Public Save() As Double             'Ahorros
Public Sol As Integer               'Variable indicadora para la generación de la solución
Public Solution() As Integer        'Vector solución
Public Sum As Double                'Variable auxiliar para el cálculo de probabilidades
Public Summary(14, 3) As Double     'Resumen de resultados
Public TabuList() As Boolean        'Lista Tabu
Public Tao() As Double              'Feromona - Parámetro propio del método
Public Time As Double               'Tiempo
Public TimeC As Single              'Tiempo máximo de una ruta
Public TimeS() As Double            'Tiempo de servicio del cliente i
Public Tused() As Double            'Tiempo consumido en una ruta (incluye transporte y servicio)
Public TusedAnt() As Single
Public Weight As Single             'peso de la feromona en el cálculo de probabilidades
Public Worst(1 To 3) As Integer           'Peor asignación (pa' quitala)
Public X() As Double                'Posición X de cada nodo
Public Y() As Double                'Posición Y de cada nodo


Type Ruta
    Ctrl As Integer
    Covered As Double
    Cused As Integer
    Tused As Double
End Type
