Attribute VB_Name = "Variables"
Public Assigned() As Boolean        'Indica si los nodos ya están asignados
Public BestCovered() As Double      'Mejor valor de la Funcion Objetivo
Public BestCused() As Integer       'Capacidad usada de cada vehículo (Mejor solución)
Public BestSolution() As Double     'Mejor Solucion
Public Capv As Integer              'Capacidad de los vehículos
Public Chosen As Integer            'Variable que indica el nodo elegido en algún paso del algoritmo
Public Covered() As Double          'Distancia recorrida por cada vehículo
Public Ctrl() As Integer            'Variable de control que indica la posición inicial y final de cada vehículo en el vector solución
Public Cused() As Integer           'Capacidad usada de cada vehículo
Public DeltaTao() As Double         'Parámetro propio del método
Public Dem() As Integer             'Demanda de cada nodo
Public Dist() As Double             'Distancia entre cada par de nodos
Public Eta() As Double              'Información heurística - parámetro propio del método
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
Public Tao() As Double              'Feromona - Parámetro propio del método
Public Time As Double               'Tiempo
Public TimeC As Single              'Tiempo máximo de una ruta
Public TimeS() As Single            'Tiempo de servicio del cliente i
Public Tused() As Single            'Tiempo consumido en una ruta (incluye transporte y servicio)
Public Weight As Single             'peso de la feromona en el cálculo de probabilidades
Public X() As Double                'Posición X de cada nodo
Public Y() As Double                'Posición Y de cada nodo

