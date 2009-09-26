Attribute VB_Name = "Variables"
Public Assigned() As Boolean        'Indica si los nodos ya están asignados
Public BestAnt() As Integer          'Mejor hormiga de cada iteración
Public BestNv As Integer            'Número de vehículos en la mejor solución
Public BestRoute() As Issues
Public BestSolution() As Integer    'Mejor Solucion
Public Capv As Integer              'Capacidad de los vehículos
Public Chosen As Integer            'Variable que indica el nodo elegido en algún paso del algoritmo
Public DeltaTao() As Double         'Parámetro propio del método
Public Dem() As Integer             'Demanda de cada nodo
Public Demand As Integer            'Demanda total (Suma)
Public DemandA As Integer           'Demanda Asignada, o en su defecto Acumulada
Public Dist() As Double             'Distancia entre cada par de nodos
Public Eta() As Double              'Información heurística - parámetro propio del método
Public Final(1 To 14) As Collect
Public First() As Double            '??? Ni idea - Si funciona le buscamos la lógica
Public Furthest As Integer          'Nodo más lejano
Public GenerationNext As Long       'Indice de generaciones
Public History() As Double          'Almace los resultados de toda la ejecución
Public K As Integer                 'Contador que representa la indexacion de cada hormiga
Public LastImprove As Long          'Última iteración en la que se mejoró la solución
Public Max As Double                'Máximo
Public MaxNv As Integer             'Número máximo de vehículos
Public MaxX As Integer              'Maxima coordenada en X (para efectos graficos)
Public MaxY As Integer              'Maxima coordenada en Y (para efectos graficos)
Public MinX As Integer              'Minima coordenada en X (para efectos graficos)
Public MinY As Integer              'Minima coordenada en Y (para efectos graficos)
Public nAnts As Integer             'Número de hormigas
Public Nearest() As Integer         'Nodo más cercano (inicio o final de la ruta) a cada nodo
Public nGen As Long                 'Número de generaciones
Public NLS As Byte                  'Número de Búsquedas Locales
Public Nodes As Integer             'Número de nodos
Public NProblem As Integer          'Número del problema
Public Nv As Integer                'Número de vehículos
Public p As Single                  'Exponente de la feromona
Public Prob As Double               'Variable auxiliar para el cálculo de probabilidades
Public Problem As String            'Problema(s) a resolver
Public Random As Double             'Variable auxiliar para el uso de números aleatorios
Public Rho As Single
Public Route() As Issues
Public RouteAnt() As Issues
Public Samples As Integer           'Indica la muestra
Public Save() As Double             'Ahorros
Public Second() As Double           'Copia de first antes de eliminar los nodos que no disminuyen la distancia promedio recorrida
Public Sol As Integer               'Variable indicadora para la generación de la solución
Public Solution() As Integer        'Vector solución
Public Sum As Double                'Variable auxiliar para el cálculo de probabilidades
Public SumFirst As Integer          'Suma del número de nodos con First diferente de cero
Public Summary() As Collect         'Resumen de resultados
Public TabuList() As Boolean        'Lista Tabu
Public Tao() As Double              'Feromona - Parámetro propio del método
Public TheBest As Double            'Mejor recorrido
Public TheWorst As Double           'Peor recorrido
Public Time As Double               'Tiempo
Public TimeC As Single              'Tiempo máximo de una ruta
Public TimeS() As Double            'Tiempo de servicio del cliente i
Public Weight As Single             'peso de la feromona en el cálculo de probabilidades
Public Worst(1 To 3) As Integer           'Peor asignación (pa' quitala)
Public X() As Double                'Posición X de cada nodo
Public Y() As Double                'Posición Y de cada nodo


Type Issues
    Ctrl As Integer
    Covered As Double
    Cused As Integer
    Tused As Double
End Type


Type Neighbor
    FV As Integer       'Fisrt Vehicle
    SV As Integer       'Second Vehicle
    NFV As Integer      'Node First Vehicle
    NSV As Integer      'Node Second Vehicle
    Covered As Double   'Covered
    Help As Integer     'Help
End Type


Type Collect
    Covered As Double
    Nv As Integer
    Time As Double
End Type
