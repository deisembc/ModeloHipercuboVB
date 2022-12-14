VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Escolher exemplo:"
      Height          =   1095
      Left            =   8520
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "exemplo 2"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "exemplo 1"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5800
      Left            =   240
      ScaleHeight     =   5800
      ScaleMode       =   0  'User
      ScaleWidth      =   7065
      TabIndex        =   1
      Top             =   480
      Width           =   7125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hipercubo"
      Height          =   735
      Left            =   8520
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mostrar na tela:"
      Height          =   1815
      Left            =   8520
      TabIndex        =   5
      Top             =   2760
      Width           =   2415
      Begin VB.CheckBox Check4 
         Caption         =   "Tempos"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Medidas de desempenho"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Calculos iniciais"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dados de entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'var global
'dimensionar cte de retorno do ResSistEq para o Hipercubo
Dim SIST_POS_DET%
Dim SIST_POS_IND_IMP%
Dim IMP_T_DADOS_ENTRADA%
Dim IMP_T_CALCULOS_INICIAIS%
Dim IMP_T_MED_DESEMP%
Dim IMP_T_TEMPOS%

Dim VDir%()

Private Sub Form_Load()
'inicializar var para ResSistEqu
  SIST_POS_DET = 0
  SIST_POS_IND_IMP = 1
  IMP_T_DADOS_ENTRADA = 0
  IMP_T_CALCULOS_INICIAIS = 0
  IMP_T_MED_DESEMP = 0
  IMP_T_TEMPOS = 0
  IMP_T_DADOS_ENTRADA = 0
  IMP_T_CALCULOS_INICIAIS = 0
  IMP_T_MED_DESEMP = 0
  IMP_T_TEMPOS = 0
End Sub

Private Sub Check1_Click()
  '** controle de impressao na tela dos dados de entrada
  If Check1.Value = 1 Then
    IMP_T_DADOS_ENTRADA = 1
  Else
    IMP_T_DADOS_ENTRADA = 0
  End If
End Sub

Private Sub Check2_Click()
  '** controle de impressao na tela dos calculos iniciais
  If Check2.Value = 1 Then
    IMP_T_CALCULOS_INICIAIS = 1
  Else
    IMP_T_CALCULOS_INICIAIS = 0
  End If
End Sub

Private Sub Check3_Click()
  '** controle de impressao na tela da medida de desempenho
  If Check3.Value = 1 Then
    IMP_T_MED_DESEMP = 1
  Else
    IMP_T_MED_DESEMP = 0
  End If
End Sub

Private Sub Check4_Click()
  '** controle de impressao na tela dos tempos de execucao do programa
  If Check4.Value = 1 Then
    IMP_T_TEMPOS = 1
  Else
    IMP_T_TEMPOS = 0
  End If
End Sub

Private Sub Command1_Click()
  
  '* Programa Principal *
  'sub para chamar hipercubo
  
  '* contagem do tempo
  HIniProg = Timer

  '** definicao constantes, variaveis e matrizes para hipercubo
  
  '* dados de entrada *
  Dim iNumServ%                     '* entrada int *
                                    '  Numero total de servidores do sistema (ou de viaturas, ou de regioes)
  Dim iNumAtom%                     '* entrada int *
                                    '  Numero total de atomos em que a regiao esta dividida
  Dim fMatrizAtom!()                '* entrada simples *
                                    '  Matriz com dados dos atomos, dimensao (iNumAtom,2)
                                    '    linhas - atomos
                                    '    col 0 = Numero do atomo (de 1 a iNumAtom)
                                    '    col 1 = Taxa de chamados do atomo (fLambda do atomo i)
  Dim iMatrizPrefDespAtom%()         '* entrada int *
                                    '  Matriz que da a lista de pref de despacho das
                                    '  viaturas para os atomos, dim (iNumAtom,iNumServ)
                                    '    linhas - atomos
                                    '    col 0 = unidade primaria
                                    '    col 1 = 1a unidade backup
                                    '    col 2 = 2a unidade backup
                                    '    ...
                                    '    col iNumServ = unidade iNumServ backup
  'mi novo
  'nao le mais fMi!
  ' dim fMi!
  'le a matriz de servidores que tem mi agora
  Dim fMatrizServ!()                 '* entrada matriz
                                     '  taxa de execucao de servicos por servidor
                                     '  linhas - servidores
                                     '  col 0 = num servidor
                                     '  col 1 = taxa mi de execucao de servicos
  Dim fMatrizTDeslAtom!()            '* entrada simples * Matriz com o tempo de deslocamento entre os atomos
  Dim fMatrizProbLocalUnidNAtomJ!()  '* entrada simples * Matriz com a probabilidade que a unid n esteja localizada no atomo j
                                     '  enquanto esta disponivel para dispacho, dim(iNumServ,iNumAtom)
  
  '* constantes *
  NUM_COL_M_ATOM% = 2                '* cte * numero de colunas da matriz atomo
  COL_ATOM_M_ATOM% = 0               '* cte * coluna onde esta o n do atomo na matriz atomo
  COL_LAMB_M_ATOM% = 1               '* cte * coluna onde esta o fLambda na matriz atomo
  COL_U_PRIM_M_PREF_DESP% = 0        '* cte * coluna onde esta a unid prim na matriz de pref de despacho
  COL_U_1BACK_M_PREF_DESP% = 1       '* cte * coluna onde esta a unid de 1o backup na matriz de pref de despacho
  'mi novo
  NUM_COL_M_SERV% = 2                '* cte * numero de colunas da matriz servidores
  COL_SERV_M_SERV% = 0               '* cte * coluna onde esta o num do servidor na matriz Servidor
  COL_MI_M_SERV% = 1                 '* cte * coluna onde esta a taxa mi do servidor na matriz Servidor
  
  
  '** Leitura dos dados de entrada do programa do arquivo hipercuboDados.txt
  'verificar se o exemplo de dados foi escolhido
  If Option1 = True Then
    Open "d:\home\deise\tese\Aplicacao\dados\hipercuboDados1.txt" For Input As #1
  Else
    If Option2 = True Then
      Open "d:\home\deise\tese\Aplicacao\dados\hipercuboDados2.txt" For Input As #1
    Else
      MsgBox "Exemplo nao escolhido"
      Exit Sub
    End If
  End If
  'Leitura numero de servidores
  Line Input #1, linha$
  iNumServ = Val(linha$)
  If IMP_T_DADOS_ENTRADA = 1 Then
    MsgBox "Numero de servidores: " + Str$(iNumServ)
  End If
  
  'Leitura numero de atomos
  Line Input #1, linha$
  iNumAtom = Val(linha$)
  If IMP_T_DADOS_ENTRADA = 1 Then
    MsgBox "Numero de atomos: " + Str$(iNumAtom%)
  End If

  'Leitura valores da Matriz Atomo
  ReDim fMatrizAtom(iNumAtom - 1, NUM_COL_M_ATOM)
  For i% = 0 To iNumAtom - 1
    For j% = 0 To NUM_COL_M_ATOM - 1
      Line Input #1, linha$
      fMatrizAtom(i%, j%) = Val(linha$)
    Next
  Next
  If IMP_T_DADOS_ENTRADA = 1 Then
    ImpTMS "fMatrizAtom: ", 0, iNumAtom - 1, 0, NUM_COL_M_ATOM - 1, fMatrizAtom
  End If
  
  'Leitura valores da iMatrizPrefDespAtom
  ReDim iMatrizPrefDespAtom%(iNumAtom - 1, iNumServ - 1)
  For i% = 0 To iNumAtom - 1
    For j% = 0 To iNumServ - 1
      Line Input #1, linha$
      iMatrizPrefDespAtom%(i%, j%) = Val(linha$)
    Next
  Next
  If IMP_T_DADOS_ENTRADA = 1 Then
    ImpTMI "iMatrizPrefDespAtom: ", 0, iNumAtom - 1, 0, iNumServ - 1, iMatrizPrefDespAtom
  End If

'mi novo
  '* nao le mais a taxa de execucao de servico fMi
  ' Line Input #1, linha$
  ' fMi! = Val(linha$)
  'If IMP_TELA = 1 Then
  '  MsgBox "Taxa de execucao de servicos: " + Str$(fMi)
  'End If
  '* e sim a matriz de taxas de execucao de servicos fMatrizServ
  'Leitura da fMatrizServ
  ReDim fMatrizServ!(iNumServ - 1, NUM_COL_M_SERV - 1)
  For i% = 0 To iNumServ - 1
    For j% = 0 To NUM_COL_M_SERV - 1
      Line Input #1, linha$
      fMatrizServ!(i%, j%) = Val(linha$)
    Next
  Next
  If IMP_T_DADOS_ENTRADA = 1 Then
    ImpTMS "fMatrizServ: ", 0, iNumServ - 1, 0, NUM_COL_M_SERV - 1, fMatrizServ
  End If
  
  'Leitura valores da fMatrizTDeslAtom
  ReDim fMatrizTDeslAtom!(iNumAtom - 1, iNumAtom - 1)
  For i% = 0 To iNumAtom - 1
    For j% = 0 To iNumAtom - 1
      Line Input #1, linha$
      fMatrizTDeslAtom(i%, j%) = Val(linha$)
    Next
  Next
  If IMP_T_DADOS_ENTRADA = 1 Then
    ImpTMS "fMatrizTDeslAtom: ", 0, iNumAtom - 1, 0, iNumAtom - 1, fMatrizTDeslAtom
  End If
  'Leitura valores da fMatrizProbLocalUnidNAtomJ
  ReDim fMatrizProbLocalUnidNAtomJ!(iNumServ% - 1, iNumAtom% - 1)
  For i% = 0 To iNumServ - 1
    For j% = 0 To iNumAtom - 1
        Line Input #1, linha$
        fMatrizProbLocalUnidNAtomJ(i%, j%) = Val(linha$)
    Next
  Next
  If IMP_T_DADOS_ENTRADA = 1 Then
    ImpTMS "fMatrizProbLocalUnidNAtomJ: ", 0, iNumServ - 1, 0, iNumAtom - 1, fMatrizProbLocalUnidNAtomJ
  End If
  Close #1

  'chamar hipercubo
  Hipercubo iNumServ%, iNumAtom%, fMatrizAtom!, iMatrizPrefDespAtom%, fMatrizServ!, _
            fMatrizTDeslAtom!, fMatrizProbLocalUnidNAtomJ!, _
            NUM_COL_M_ATOM%, COL_ATOM_M_ATOM%, COL_LAMB_M_ATOM%, COL_U_PRIM_M_PREF_DESP%, _
            COL_U_1BACK_M_PREF_DESP%, NUM_COL_M_SERV%, COL_SERV_M_SERV%, COL_MI_M_SERV%
  
  HFinProg = Timer
  TTotProg = (HFinProg - HIniProg)
  If IMP_T_TEMPOS = 1 Then
    MsgBox "Tempo total (s) : " + Str$(TTotProg)
  End If
  
  MsgBox "FIM"

End Sub

Sub Hipercubo(iNumServ%, iNumAtom%, fMatrizAtom!(), iMatrizPrefDespAtom%(), fMatrizServ!(), _
              fMatrizTDeslAtom!(), fMatrizProbLocalUnidNAtomJ!(), _
              NUM_COL_M_ATOM%, COL_ATOM_M_ATOM%, COL_LAMB_M_ATOM%, COL_U_PRIM_M_PREF_DESP%, _
              COL_U_1BACK_M_PREF_DESP%, NUM_COL_M_SERV%, COL_SERV_M_SERV%, COL_MI_M_SERV%)

  '* constante
  MASCARA_INT% = 2                  '* cte * mascara para comparacao de vetores
  
  '* variaveis *
  Dim iNumVertHiper%         '* calc * Numero de vertices do hipercubo, iNumServ ate 14, pois e inteiro
  Dim iVertice1%             '* var aux * variacao do vertice do hipercubo, em decimal
  Dim iVertice2%             '* var aux * variacao do vertice do hipercubo, em decimal
  Dim iVetoresIguais%        '* var aux * compara os vetores binarios dos vertices do hipercubo
  Dim iPot%                  '* valor maximo = 32 *??????????????ver depois
  Dim fLambda!               '* Calc - Valor das taxas de chamadas de toda a regiao
  Dim fMi!                   '* Calc - Taxa de execucao de servicos
  Dim fP0!                   '* Calc - Prob de nenhum servidor ocupado
  Dim fSPn!                  '* Calc - Prob de todos os servidores ocupados
  Dim fPq!                   '* Calc - Prob de haver uma fila de comprimento positivo
  Dim fAux!                  '* var aux * var para troca de 2 valores de posicao
  Dim fTaxa!                 '* var aux * var para calculo das taxas de transicao de estado
  Dim fPlQ!                  '* Calc - Valor P´q = prob de que 1 chamada fique na fila
  Dim fFracaoInterArea!      '* Calc - b1) Fracao dentre todos os despachos que sao interareas de cobertura
  Dim fSoma!                 '* var aux * ver se contas ok
  Dim fNumerador!            '* var aux * acumular numerador
  Dim fDenominador!          '* var aux * acumular denominador
  Dim fTMedDeslChamFila!     '* Calc - c2) Valor TbarraQ = tempo medio de deslocamento para chamados da fila
  Dim fTMedGlobalDesl!       '* Calc - c3) Valor Tbarra = Tempo medio global de deslocamento
  Dim iVetorVert1%()         '* var aux * variacao do vertice do hipercubo, em binario
  Dim iVetorVert2%()         '* var aux * variacao do vertice do hipercubo, em binario
  Dim fVetorLambda!()        '* Calc - valores das taxas de chamadas por area de cobertura */
  Dim fVetorProbEst!()       '* Inicialmente tem valores do termo ind do sist eq de equilibrio do sistema
                             '* Ao final tem Calc - Prob do sistema estar no estado {ijk}, sol do sist eq, (linha em decimal = ijk em binario) */
  Dim fVetorOcupUnidN!()     '* Calc - a) Valor ro = carga de trabalho da unid, quant do t que esta ocupado */
  Dim fVetorFracao2!()       '* Calc - b) Frequencias de despachos interatomos, fracao dentre todos os
                             '* despachos que designam n ao atomo j para atender um chamado que
                             '* estava na fila de espera, cada linha refere-se a 1 atomo */
  Dim fVetorFracaoInterN!()  '* Calc - b2) Fracao dos todos os despachos da unidade n que sao interareas
                             '* de cobertura (somente os backups), cada linha refere-se a 1 unid */
  Dim fVetorFracaoInterNaoN!() '* Calc - b3) Fracao dos chamados da area N que sao atendidas por outra unid que nao a N */
  Dim fVetorTMedDeslAtomJ!()   '* Calc - c4) Tempo medio de deslocamento ate o atomo j */
  Dim fVetorTMedDeslAreaN!()   '* Calc - c5) Tempo medio de deslocamento para uma dada area de cobertura primaria */
  Dim fVetorTMedDeslUnidN!()   '/* Calc - c6) Expressao aproximada para Tempo medio de deslocamento da unidade n */
  Dim fMatrizTTE!()            '/* Calc - matriz de taxa de transicao de estado, (estado em binario=numero do vert em decimal) */
  Dim fMatrizFracao1!()        '* Calc - b) Frequencias de despachos interatomos, fracao dentre todos os
                              '* despachos que designam a unid n ao atomo j para atender um chamado que
                              '* nao entrou na fila de espera */
  Dim fMatrizTMedDeslNJ!()     '* Calc - c1) Tempo medio requerido a unid n quando disponivel,
                              '*            para deslocar-se ate o atomo j

  Dim iRetResSistEq%          '* retorno do resultado do sistema de equacoes


  
  


  '** Obter valores para o Sistema de equacoes

  '** Calcular fLambda = taxa total de chamados da regiao toda
  fLambda! = 0
  For i% = 0 To iNumAtom - 1
    fLambda! = fLambda! + fMatrizAtom!(i%, COL_LAMB_M_ATOM)
  Next
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    MsgBox "Taxa de chamadas de toda regiao (fLambda) = " + Format(fLambda!, "#0.00000")
  End If
  
  
'mi novo
  '** Calcular fMi = taxa total de execucao de servicos
  fMi! = 0
  For i% = 0 To iNumServ - 1
    fMi = fMi + fMatrizServ(i%, COL_MI_M_SERV)
  Next
  fMi = fMi / iNumServ
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    MsgBox "Taxa de execucao de servicos: " + Str$(fMi)
  End If
  
  
  
  '** Calculo das probabilidades P
  '* fP0 = probabilidade de ter 0 servidores ocupados
  fP0! = 0
  For i% = 0 To iNumServ - 1
    fP0! = fP0! + ((fLambda! / fMi!) ^ i%) / fatorial(i%)
  Next
  fP0! = fP0! + (((fLambda! / fMi!) ^ iNumServ%) / fatorial(iNumServ%)) * ((iNumServ% * fMi!) / (iNumServ% * fMi! - fLambda!))
  fP0! = 1 / fP0!
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    MsgBox "Probabilidade de ter 0 servidores ocupados (fP0) = " + Format(fP0!, "#0.00000")
  End If
  '* Calculo da soma das probabilidade de ter n servidores ocupados, 0<=n<=NumServ
  fSPn! = fP0!
  For i% = 1 To iNumServ
    fSPn! = fSPn! + (((fLambda! / fMi!) ^ i%) / fatorial(i%)) * fP0!
  Next
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    MsgBox "Soma das probabilidade de ter n servidores ocupados, 0<=n<=NumServ (fSPn) = " + Str$(fSPn!)
  End If
  '* fPq = probabilidade de haver uma fila de comprimento positivo
  fPq! = 1 - fSPn!
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    MsgBox "Probabilidade de haver uma fila de comprimento positivo (Pq=1-fSPn) = " + Str$(fPq!)
  End If
  
  
  '** Calcular vetor fLambda, guarda os valores das taxas de chamadas por viatura/servidor
  ReDim fVetorLambda!(iNumServ - 1)
  'pesquisar onde aparece a unidade i na MatrizPrefDesp, como unidade primaria
  For i% = 0 To iNumServ - 1 'procurar o servidor i dentro da matriz
    For j% = 0 To iNumAtom - 1 'variar os atomos
      'o atomo e atendido pelo servidor i? */
      If iMatrizPrefDespAtom%(j%, COL_U_PRIM_M_PREF_DESP) = i% Then
        fVetorLambda(i%) = fVetorLambda(i%) + fMatrizAtom(j%, COL_LAMB_M_ATOM)
      End If
    Next
  Next
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    ImpTVS "Taxa de chamados por area de cobertura (fVetorLambda)= ", 0, iNumServ - 1, fVetorLambda!
  End If
  
  
  '** Numero de vertices do hipercubo, numerados de 0 a iNumVertHiper-1
  'para obter os estados, basta transf o numero decimal do vertice em binario */
  iNumVertHiper% = 2 ^ iNumServ%
  
  
  '** Obter matriz de taxa de transicao de estado,  fMatrizTTE
  'estado(000)= nenhum servidor atendendo, estado(100)= servidor 1 atendendo
  ReDim fMatrizTTE!(iNumVertHiper - 1, iNumVertHiper - 1) 'matriz: 0 ao iNumVertHiper-1
  
  '* Variar os vertices do hipercubo
  '* Caso especial: iVertice1% = 0 = estado(000)
  'fTaxa = fLambda da regiao atendida pelo servidor j
  iVertice1% = 0 'estado (000)
  iPot% = iNumServ 'controle da potencia, pot=2->serv=2^2=4->estado(100)
  'regioes atendidas pelo servidor j=1,2,3; estado=(100),(010),(001); vertice2=4,2,1
  For j% = 0 To iNumServ - 1
    iPot% = iPot% - 1 'pot=2,1,0
    iVertice2 = 2 ^ iPot
    'taxa de transicao ascendente
    fMatrizTTE(iVertice1, iVertice2) = fVetorLambda(j%) 'vertices: 0->4, 0->2, 0->1
    'taxa de transicao descendente
'mi novo
    fMatrizTTE(iVertice2, iVertice1) = fMatrizServ(iPot, COL_MI_M_SERV)
    
  Next
  
  
  '* Demais casos vertices i>0, a variacao da numeracao dos vertices eh de 0 a iNumVertHiper-1 */
  For i% = 1 To iNumVertHiper% - 2
    iVertice1% = i%
    'converter iVertice1 em binario => iVetorVert1
    ReDim iVetorVert1%(iNumServ - 1) 'zerar
    dec2Bin iVertice1%, iNumServ - 1, iVetorVert1%
    iPot = iNumServ%
    'Variar servidores que podem comecar a trabalhar a partir do vertice i
    For j = 0 To iNumServ - 1
    'If j = 7 Then Stop
      iPot = iPot - 1
      'servidor j esta livre para comecar a trabalhar?
      If iVetorVert1%(j) <> 1 Then '(001)
      'sim
        iVertice2% = iVertice1% + 2 ^ iPot%
        'V2 nao e o ultimo vertice ?*/
        If iVertice2% <> iNumVertHiper - 1 Then
        'sim, nao e o ultimo vertice do hiper
          fTaxa = fVetorLambda(j%) 'fLambda da regiao primaria atendida pelo servidor j
          'variar os servidores k que ja estao atendendo, pois o chamado pode ter vindo da regiao na qual o j e backup
          For k = 0 To iNumServ - 1
            'servidor k esta atendendo?
            If iVetorVert1(k) = 1 Then
            'sim
              'obter fLambda dos atomos l da regiao do servidor k no qual j é 1o backup
              For l% = 0 To iNumAtom - 1
                'servidor j é 1o backup do atomo l da regiao k? */
                If ((iMatrizPrefDespAtom(l, COL_U_1BACK_M_PREF_DESP) = j) And (iMatrizPrefDespAtom(l, COL_U_PRIM_M_PREF_DESP) = k)) Then
                  fTaxa = fTaxa + fMatrizAtom(l, COL_LAMB_M_ATOM)
                End If
              Next
            End If
          Next
          'fTaxa de transicao ascendente
          fMatrizTTE(iVertice1, iVertice2) = fTaxa
          'fTaxa de transicao descendente
'mi novo
          fMatrizTTE(iVertice2, iVertice1) = fMatrizServ(iPot, COL_MI_M_SERV)
        Else
        'sim, é o ultimo vertice - Caso especial
          'fTaxa de transicao ascendente
          fMatrizTTE(iVertice1, iVertice2) = fLambda 'recebe taxa total
          'taxa de transicao descendente
'mi novo
          fMatrizTTE(iVertice2, iVertice1) = fMatrizServ(iPot, COL_MI_M_SERV)
        End If
      End If
    Next
  Next
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    ImpTMS "Matriz de taxa de transicao de estados (fMatrizTTE) : ", 0, iNumVertHiper - 1, 0, iNumVertHiper - 1, fMatrizTTE
  End If
  

  '** Obter os coeficientes das Equacoes de equilibrio do sistema
  'Fornecem as probabilidades PIijk do sistema estar no estado {ijk}
  'Resposta do sistema esta no fVetorProbEst
  'Fluxo de saida = fluxo de entrada em cada vertice */
  'Fluxo de saida do vertice i é dado pela soma dos elementos da
  '      linha da fMatrizTTE - sera guardado na propria fMatrizTTE(i,i) com sinal -
  'Fluxo de entrada dos vertices i para o vertice j é dado
  '      pelas colunas da fMatrizTTE multiplicada pelo Xi correspondente
  '      os coeficientes permanecem os mesmos da matriz original
  'logo: A matriz de coeficientes do sistema é a transposta da fMatrizTTE
  '      e que contem na diagonal principal a soma das linhas da fMatrizTTE
  'Calculo da diagonal principal da fMatrizTTE, fluxo de saida */
  For i% = 0 To iNumVertHiper - 1
    For j% = 0 To iNumVertHiper - 1
      If i% <> j% Then
        fMatrizTTE(i%, i%) = fMatrizTTE(i%, i%) + fMatrizTTE(i%, j%)
      End If
    Next
    fMatrizTTE(i%, i%) = fMatrizTTE(i%, i%) * -1
  Next
  'FAZER DEPOIS: eh so trocar em cima iVertice1 com iVertice2, ja dara a transposta!!
  'E ao inves de somar linhas, somar colunas */
  '* Obter a transposta da fMatrizTTE
  For i = 0 To iNumVertHiper - 2
    For j = i + 1 To iNumVertHiper - 1
      fAux = fMatrizTTE(i, j)
      fMatrizTTE(i, j) = fMatrizTTE(j, i)
      fMatrizTTE(j, i) = fAux
    Next
  Next
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    ImpTMS "Matriz dos coeficientes das equacoes de equilibrio do sistema (fMatrizTTE):", 0, iNumVertHiper - 1, 0, iNumVertHiper - 1, fMatrizTTE
  End If
  '* Substituir a ultima linha da fMatrizTTE pela equacao que representa
  'as probabilidades de ocupacao dos estados para um sistema com fila de capacidade infinita
  For j = 0 To iNumVertHiper - 1
    fMatrizTTE(iNumVertHiper - 1, j) = 1
  Next
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    ImpTMS "Matriz dos coeficientes das equacoes de equilibrio do sistema, subst a ultima eq pela das prob de ocupacao dos estados para um sistema com fila (fMatrizTTE):", 0, iNumVertHiper - 1, 0, iNumVertHiper - 1, fMatrizTTE
  End If
  
  
  'obter vetor dos termos independentes
  ReDim fVetorProbEst!(iNumVertHiper - 1)
  'atualizacao do fVetorProbEst na ultima posicao por 1-fPq
  fVetorProbEst(iNumVertHiper - 1) = 1 - fPq
  If IMP_T_CALCULOS_INICIAIS = 1 Then
    ImpTVS "Vetor do termo independente das equacoes de equilibrio do sistema (fVetorProbEst): ", 0, iNumVertHiper - 1, fVetorProbEst
  End If

  '** Solucionar o sistema de equacoes de equilibrio
  '* colocar o retorno da funcao na var iRetResSistEq
  '* resposta do sistema esta no proprio vetor fVetorProbEst
  
  ReDim VDir%(iNumVertHiper% - 1)

  HIniResSistEq = Timer
  iRetResSistEq = resSistEq(iNumVertHiper%, fVetorProbEst!, fMatrizTTE!)
  HFinResSistEq = Timer
  TResSistEq = HFinResSistEq - HIniResSistEq
  If IMP_T_TEMPOS = 1 Then
    MsgBox "Tempo resolucao do sistema de equacoes (s) : " + Str$(TResSistEq)
  End If
  
  'sistema possivel?
  Select Case iRetResSistEq
  Case 0 'sistema possivel determinado
    If IMP_T_CALCULOS_INICIAIS = 1 Then
      ImpTVS "Probabilidades do sistema estar no estado {ijk}(em binario) (fVetorProbEst):", 0, iNumVertHiper - 1, fVetorProbEst
    End If
  Case 1 'sistema possivel determinado
    MsgBox "Sistema Possivel Indeterminado ou Impossivel."
    Stop
  'Case Else 'se nao for nenhum dos valores acima
  End Select
 
 
'** Calculo das medidas de desempenho do sistema
  'impTela("Calculo das medidas de desempenho do sistema");
  '* a) Carga de trabalho: fornecida pelo fVetorOcupUnidN
  '*    indica o quanto de tempo que a unidade n esta ocupada */
  '* FAZER DEPOIS: a soma do fVetorOcupUnidN tem que dar fLambda, da para verificar se esta ok! */
  'dimensionar o fVetorOcupUnidN
  ReDim fVetorOcupUnidN!(iNumServ% - 1)
  '* pequisar as area de cobertura = regioes
  For i = 0 To iNumServ - 1
    fVetorOcupUnidN(i) = fPq
    'pesquisar os vertices
    For j = 0 To iNumVertHiper - 1
      'converter o vertice j em bin
      ReDim iVetorVert1(iNumServ - 1) 'zerar
      dec2Bin j, iNumServ - 1, iVetorVert1
      'Posicao i do iVetorVert1 e 1?
      If iVetorVert1(i) = 1 Then
      'sim
        'incluir o valor de PI referente ao vertice j */
        fVetorOcupUnidN(i) = fVetorOcupUnidN(i) + fVetorProbEst(j)
      End If
    Next
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTVS "a) Carga de trabalho das unidades (fVetorOcupUnidN): ", 0, iNumServ - 1, fVetorOcupUnidN
  End If
  

  '** b) Frequencias de despachos interatomos
  'Calculo de f1
  'Resposta para f1 esta na fMatrizFracao1, dim[iNumAtom][iNumServ]
  '   col 0        = f1 unid0 atomoj
  '   col 1        = f1 unid1 atomoj
  '   col 2        = f1 unid2 atomoj
  '   ...
  'dimensionar matriz fMatrizFracao1
  ReDim fMatrizFracao1(iNumAtom - 1, iNumServ - 1)
  'Determinar o conj de todos os estados nos quais a unidade n pode
  ' ser designada para atender ao atomo j - Enj
  'Determinar 1 estado de cada vez, usando o iVetorVert1,
  ' na posicao n sera 0 - unid n esta livre para atender atom j
  ' na matrizPrefDesp verificar na linha do atom j, as unid de atend
  ' que tem pref de despacho antes da n, estas devem estar ocupadas,
  ' logo tem que ter o valor 1 na pos corresp do iVetorVert1
  'Dimensionar memoria para o iVetorVert2
  ReDim iVetorVert2(iNumServ - 1)
  'variar unidade n
  For n = 0 To iNumServ - 1
    'variar atomo j
    For j = 0 To iNumAtom - 1
      'setar o iVetorVert1 com valor 2
      SetarValorVetorInt iNumServ - 1, iVetorVert1, MASCARA_INT%
      'na posicao n do iVetorVert1 tem que ter 0
      iVetorVert1(n) = 0
      'variar a coluna do atomo j na iMatrizPrefDespAtom ate encontrar n
      For i = 0 To iNumServ - 1
        'achei a posicao n?
        If iMatrizPrefDespAtom(j, i) = n Then
        'sim
          'sair da procura
          Exit For
        Else
        'nao
          'atualizar iVetorVert1 na posicao da unid ocupada
          iVetorVert1(iMatrizPrefDespAtom(j, i)) = 1
        End If
      Next
      'variar todos os vertices do hipercubo, comparar com a mascara que esta no iVetorVert1
      '  pegar os que tem 0 na posicao n e 1 na posicao das demais
      '  unidades que tem pref de despacho antes da unidade n
      For i = 0 To iNumVertHiper - 1
        iVetoresIguais = 1
        'transformar i em bin
        ReDim iVetorVert2(iNumServ - 1)
        dec2Bin i, iNumServ - 1, iVetorVert2
        'verificar se os dois sao iguais, exceto pelas posicoes "2"
        For k = 0 To iNumServ - 1
          'posicoes nao sao iguais?
          If ((iVetorVert1(k) <> iVetorVert2(k)) And (iVetorVert1(k) <> MASCARA_INT)) Then
          'sim
            iVetoresIguais = 0 'Vetores sao diferentes
            Exit For 'TESTAR
          End If
        Next
        If iVetoresIguais = 1 Then
        'sim
          fMatrizFracao1(j, n) = fMatrizFracao1(j, n) + fVetorProbEst(i) '= prob PI do vertice i
        End If
      Next 'fim for i
      'multiplicar por lambdaUnidadeJ/fLambda
      fMatrizFracao1(j, n) = fMatrizFracao1(j, n) * (fMatrizAtom(j, COL_LAMB_M_ATOM) / fLambda)
    Next
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTMS "b) f1jn - Fracao dentre todos os despachos que designam a unid n ao atomo j para atender um chamado que nao entrou na fila de espera  (fMatrizFracao1): ", 0, iNumAtom - 1, 0, iNumServ - 1, fMatrizFracao1
  End If

  '** Calculo de f2 - fracao dentre todos os despachos que designam a n
  '   ao atomo j para atender um chamado que estava na fila de espera
  '** Resposta para f2 esta na fVetorFracao2, dim[iNumAtom]
  '   col 0 = f2 atomoj
  'redimensionar fVetorFracao2
  ReDim fVetorFracao2(iNumAtom - 1)
  'calculo de P´q (fPlQ)
  fPlQ = fPq + fVetorProbEst(iNumVertHiper - 1)
  'obter f2
  For i = 0 To iNumAtom - 1
    fVetorFracao2(i) = ((fMatrizAtom(i, COL_LAMB_M_ATOM) * fPlQ) / (fLambda * iNumServ))
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTVS "f2jn - Fracao dentre todos os despachos que designam as unidades ao atomo j para atender um chamado que estava na fila de espera (fVetorFracao2): ", 0, iNumAtom - 1, fVetorFracao2
  End If
  
  'Mostrar na tela valor de fnj=f1nj+f2nj, somente para conferir valores
  mens$ = "fnj - Fracao dentre todos os despachos que designam a unid n ao atomo j (Matriz fnj=f1nj+f2nj): " + Chr$(13)
  For i% = 0 To iNumAtom - 1
    mens$ = mens$ + Format$(i%, "00")
    For j% = 0 To iNumServ - 1
      mens$ = mens$ + Format$(fMatrizFracao1(i, j) + fVetorFracao2(i%), " 00.00000") 'antes era chr$(10)+chr$(13)
    Next
    mens$ = mens$ + Chr$(13)
  Next
  If IMP_T_MED_DESEMP = 1 Then
    MsgBox mens$
  End If
    
  'Verificar se as contas estao batendo!!!
  'testar se a soma de todos os fnj = 1?
  fSoma = 0
  For i = 0 To iNumAtom - 1
    For j = 0 To iNumServ - 1
      fSoma = fSoma + fMatrizFracao1(i, j) + fVetorFracao2(i)
    Next
  Next
  If IMP_T_MED_DESEMP = 1 Then
    MsgBox "Verificacao: soma de todos os elementos da matriz Fracao 1 com vetor Fracao 2 = " + Str$(fSoma)
  End If

  '** b1) Fracao dentre todos os despachos que sao interareas de cobertura
  '*      Obter fFracaoInterArea
  '*      deixar de fora as unid primarias
  '** b2) Fracao dos todos os despachos da unidade n que sao interareas de cobertura
  '*      Obter fVetorFracaoInterN
  '* Calcular os dois juntos */
  '* redimensionar fVetorFracaoInterN
  ReDim fVetorFracaoInterN(iNumServ - 1)
  fFracaoInterArea = 0
  'variar o servidor n
  For i = 0 To iNumServ - 1
    fNumerador = 0
    fDenominador = 0
    'percorrer os atomos
    For j = 0 To iNumAtom - 1
      'unidade i e primaria do atomo j?
      If iMatrizPrefDespAtom(j, COL_U_PRIM_M_PREF_DESP) <> i Then
        'nao
        'considerar no calculo
        fNumerador = fNumerador + (fMatrizFracao1(j, i) + fVetorFracao2(j))
      End If
      fDenominador = fDenominador + (fMatrizFracao1(j, i) + fVetorFracao2(j))
    Next
    fVetorFracaoInterN(i) = fNumerador / fDenominador
    fFracaoInterArea = fFracaoInterArea + fNumerador
  Next
  If IMP_T_MED_DESEMP = 1 Then
    MsgBox "b) Frequencias de despachos interatomos" + Chr$(13) + "b1) Fracao dentre todos os despachos que sao interareas de cobertura (fFracaoInterArea): " + Format(fFracaoInterArea, " 00.00000")
    ImpTVS "b) Frequencias de despachos interatomos - b2) Fracao dos despachos de n que sao interarea de cobertura (fVetorFracaoInterN): ", 0, iNumServ - 1, fVetorFracaoInterN
  End If
  
  '* b3) Fracao dos chamados da area N que sao atendidas por outra unid que nao a N
  '*     Obter fVetorFracaoInterNaoN
  '* redimensionar fVetorFracaoInterNaoN
  ReDim fVetorFracaoInterNaoN(iNumServ - 1)
  'variar a area de cobertura i
  For i = 0 To iNumServ - 1
    fNumerador = 0
    fDenominador = 0
    'percorrer os atomos
    For j = 0 To iNumAtom - 1
      'j pertence a area de cobertura de i?
      If iMatrizPrefDespAtom(j, COL_U_PRIM_M_PREF_DESP) = i Then
        'sim
        'variar o servidor k
        For k = 0 To iNumServ - 1
          'servidor k = area i?
          If k <> i Then
            'nao
            fNumerador = fNumerador + (fMatrizFracao1(j, k) + fVetorFracao2(j))
          End If
          fDenominador = fDenominador + (fMatrizFracao1(j, k) + fVetorFracao2(j))
        Next
      End If
    Next
    fVetorFracaoInterNaoN(i) = fNumerador / fDenominador
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTVS "b) Frequencias de despachos interatomos - b3) Fracao dos chamados da area n que sao atendidas por outra unid que nao a n (fVetorFracaoInterNaoN): ", 0, iNumServ - 1, fVetorFracaoInterNaoN
  End If
 
  
  'c) Tempos de deslocamentos
  'c1) tnj = tempo medio requerido a unid n quando disponivel, para deslocar-se ate o atomo j
  'redimensionar fMatrizTMedDeslNJ
  ReDim fMatrizTMedDeslNJ(iNumServ - 1, iNumAtom - 1)
  'multiplicar as matrizes fMatrizProbLocalUnidNAtomJ e fMatrizTDeslAtom */
  MultiplicarMatrizes iNumServ% - 1, iNumAtom% - 1, fMatrizProbLocalUnidNAtomJ, iNumAtom% - 1, iNumAtom% - 1, fMatrizTDeslAtom, fMatrizTMedDeslNJ
  If IMP_T_MED_DESEMP = 1 Then
    ImpTMS "c) Tempos de deslocamentos - c1) Tempo medio requerido a unidade n (qdo disp) para deslocar-se ate o atomo j (fMatrizTMedDeslNJ): ", 0, iNumServ - 1, 0, iNumAtom - 1, fMatrizTMedDeslNJ
  End If
  

  '/* c2) Tempo medio de deslocamentos para chamados da fila
  fTMedDeslChamFila = 0
  For i = 0 To iNumAtom - 1
    For j = 0 To iNumAtom - 1
      fTMedDeslChamFila = fTMedDeslChamFila + (((fMatrizAtom(i, COL_LAMB_M_ATOM) * fMatrizAtom(j, COL_LAMB_M_ATOM)) / _
                          (fLambda) ^ 2) * fMatrizTDeslAtom(i, j))
    Next
  Next
  If IMP_T_MED_DESEMP = 1 Then
    MsgBox "c) Tempos de deslocamentos - c2) Tempo medio de deslocamentos para chamados da fila (fTMedDeslChamFila) = " + Str$(fTMedDeslChamFila)
  End If


  'c3) fTMedGlobalDesl = Tempo medio global de deslocamento
  fTMedGlobalDesl = 0
  For n = 0 To iNumServ - 1
    For j = 0 To iNumAtom - 1
      fTMedGlobalDesl = fTMedGlobalDesl + (fMatrizFracao1(j, n) * fMatrizTMedDeslNJ(n, j))
    Next
  Next
  fTMedGlobalDesl = fTMedGlobalDesl + fPlQ * fTMedDeslChamFila
  If IMP_T_MED_DESEMP = 1 Then
    MsgBox "c) Tempos de deslocamentos - c3) Tempo medio global de deslocamento (fTMedGlobalDesl) = " + Str$(fTMedGlobalDesl)
  End If
  
  
  'c4) Tempo medio de deslocamento ate o atomo j, fVetorTMedDeslAtomJ
  'redimensionar fVetorTMedDeslAtomJ
  ReDim fVetorTMedDeslAtomJ(iNumAtom - 1)
  For j = 0 To iNumAtom - 1
    fNumerador = 0
    fDenominador = 0
    'calculo da 1a parcela
    For n = 0 To iNumServ - 1
      fNumerador = fNumerador + (fMatrizFracao1(j, n) * fMatrizTMedDeslNJ(n, j))
      fDenominador = fDenominador + fMatrizFracao1(j, n)
    Next
    fVetorTMedDeslAtomJ(j) = (fNumerador / fDenominador) * (1 - fPlQ)
    'calculo da 2a parcela
    For i = 0 To iNumAtom - 1
      fVetorTMedDeslAtomJ(j) = fVetorTMedDeslAtomJ(j) + _
         ((fMatrizAtom(i, COL_LAMB_M_ATOM) / fLambda) * fMatrizTDeslAtom(i, j) * _
         fPlQ)
    Next
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTVS "c) Tempos de deslocamentos - c4) Tempo medio de deslocamento ate o atomo j (fVetorTMedDeslAtomJ):", 0, iNumAtom - 1, fVetorTMedDeslAtomJ
  End If


  '/* c5)  = Tempo medio de deslocamento para uma dada area de cobertura primaria
  '/* Calculo de fVetorTMedDeslAreaN = tempo medio de deslocamento ate os atomos da area de cobertura primaria N */
  '/* redimensionar fVetorTMedDeslAreaN */
  ReDim fVetorTMedDeslAreaN(iNumServ - 1)
  '/* variar as areas de cobertura */
  For n = 0 To iNumServ - 1
    fNumerador = 0
    fDenominador = 0
    ' calculo da 1a parcela
    ' variar os atomos j
    For j = 0 To iNumAtom - 1
      'j pertence a area de cobertura primaria de n?
      If iMatrizPrefDespAtom(j, COL_U_PRIM_M_PREF_DESP) = n Then
        'sim
        For m = 0 To iNumServ - 1
          fNumerador = fNumerador + (fMatrizFracao1(j, m) * fMatrizTMedDeslNJ(m, j))
          fDenominador = fDenominador + fMatrizFracao1(j, m)
        Next
      End If
    Next
    fVetorTMedDeslAreaN(n) = ((fNumerador / fDenominador) * (1 - fPlQ))
    ' calculo da 2a parcela
    fNumerador = 0
    fDenominador = 0
    'variar os atomos k
    For k = 0 To iNumAtom - 1
      'k pertence a area de cobertura primaria de n? */
      If iMatrizPrefDespAtom(k, COL_U_PRIM_M_PREF_DESP) = n Then
        'sim
        For j = 0 To iNumAtom - 1
          fNumerador = fNumerador + ((fMatrizAtom(j, COL_LAMB_M_ATOM) * _
                                   fMatrizAtom(k, COL_LAMB_M_ATOM) * fMatrizTDeslAtom(j, k)) / (fLambda * fLambda))
        Next
        fDenominador = fDenominador + (fMatrizAtom(k, COL_LAMB_M_ATOM) / fLambda)
      End If
    Next
    fVetorTMedDeslAreaN(n) = fVetorTMedDeslAreaN(n) + ((fNumerador / fDenominador) * (fPlQ))
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTVS "c) Tempos de deslocamentos - c5) Tempo medio de deslocamento para a area de cobertura primaria (fVetorTMedDeslAreaN):", 0, iNumServ - 1, fVetorTMedDeslAreaN
  End If

  '/* c6)  = Expressao aproximada para Tempo medio de deslocamento da unidade n
  '/* Calculo de TMedUnidN = tempo medio de deslocamento da unidade N */
  ' redimensionar fVetorTMedDeslUnidN
  ReDim fVetorTMedDeslUnidN(iNumServ - 1)
  'variar as areas de cobertura
  For n = 0 To iNumServ - 1
    fNumerador = 0
    fDenominador = 0
    'variar os atomos j
    For j = 0 To iNumAtom - 1
      fNumerador = fNumerador + (fMatrizFracao1(j, n) * fMatrizTMedDeslNJ(n, j))
      fDenominador = fDenominador + fMatrizFracao1(j, n)
    Next
    fVetorTMedDeslUnidN(n) = ((fNumerador + (fTMedDeslChamFila * fPlQ / iNumServ)) / _
                              (fDenominador + (fPlQ / iNumServ)))
  Next
  If IMP_T_MED_DESEMP = 1 Then
    ImpTVS "c) Tempos de deslocamentos - c6) Tempo medio aproximado de deslocamento da unidade N (fVetorTMedDeslUnidN):", 0, iNumServ - 1, fVetorTMedDeslUnidN
  End If
 
 





End Sub



Static Sub Bina(ba As Integer, bsaida As String)
'somente ate valores ate 32768 = 2^15
  bi = ba
  bb = ""
  Do While bi > 0
    bq = Fix(bi / 2)
    br% = bi - bq * 2
    bi = bq
    bb = Mid$(Str$(br%), 2) + bb
  Loop
  bsaida = bb
End Sub

Function fatorial(Numero%) 'muda valores da variavel passada
'* calcula fatorial
  
  Dim fator!
  Dim i%

  fator! = 1
  
  If Numero% < 0 Then
  'erro, nao existe fatorial negativo
    Stop
  Else
    For i% = 2 To Numero%
      fator! = fator! * i%
    Next
  End If

  fatorial = fator!

End Function

Sub dec2Bin(Numero%, NumCol%, Vetor%())
  i% = NumCol%
  a% = Numero%
  Do While a% > 0
    bq = Fix(a% / 2)
    br% = a% - bq * 2
    a% = bq
    'atualizar vetor
    Vetor%(i%) = br
    i% = i% - 1
  Loop
End Sub

Static Function resSistEq(NumEq%, VetorB!(), MatrizA!())
 '* Programa para solucionar um sistema de equacoes lineares
 '* Metodo de Gauss, pivoteamento completo
 '* Ax=b => LUx=b; Ux=y e Ly=b

  '* Parametros de entrada
  '*   NumEq   = Quantidade de equacoes do sistema
  '*   VetorB  = Vetor dos termos independentes do sistema.
  '*             Em caso de sistema possivel determinado,
  '*             ao final da funcao contem a solucao x
  '*   MatrizA = Matriz dos coeficientes do sistema, usa a linha zero
   
  '*
  '* Retorno
  '*   iTipoSist (valores que pode assumir estao definidos no resSistEq.h)
  '*     -1 = ERRO_MEM         = Erro na alocacao de memoria
  '*      0 = SIST_POS_DET     = Sistema Possivel Determinado
  '*      1 = SIST_POS_IND_IMP = Sistema Possivel Indeterminado ou Impossivel
  '*


  'definicao variaveis
  Dim Lin%
  Dim Col%
  Dim k%
  Dim PosLinMaiorEle%
  Dim PosColMaiorEle%
  Dim AuxI%
  Dim AuxS!
  Dim VetorTrocaColX%()
  Dim TipoSist%
  TipoSist% = SIST_POS_DET
  
  
  'inicializar vetor que guarda as trocas de colunas do sistema
  ReDim VetorTrocaColX%(NumEq%)
  For Lin% = 0 To NumEq - 1
    VetorTrocaColX%(Lin%) = Lin%
  Next



TipoSol = 0

If TipoSol = 0 Then
'gauss completo
  
  
  'Decomposicao LU
  For k% = 0 To NumEq% - 1
    'procurar maior elemento a partir da linha k e da coluna k
    AuxS! = Abs(MatrizA!(k%, k%))
    PosLinMaiorEle% = k%
    PosColMaiorEle% = k%
    For Lin% = k% To NumEq% - 1
      For Col% = k% To NumEq% - 1
        If (Abs(MatrizA!(Lin, Col)) > Abs(AuxS!)) Then
          AuxS! = Abs(MatrizA!(Lin%, Col%))
          PosLinMaiorEle% = Lin%
          PosColMaiorEle% = Col%
        End If
      Next
    Next

    'pivo nulo?
    If (MatrizA!(PosLinMaiorEle%, PosColMaiorEle%) = 0) Then
      'sim
      TipoSist = SIST_POS_IND_IMP
      Stop 'sistema possivel indeterminado
    End If

    'pivo esta na linha k?
    If (k% <> PosLinMaiorEle) Then
      'nao, entao trocar linha k com a linha iPosLinMaiorEle
      For Col% = 0 To NumEq - 1
        AuxS! = MatrizA(k%, Col%)
        MatrizA(k%, Col%) = MatrizA(PosLinMaiorEle%, Col%)
        MatrizA(PosLinMaiorEle%, Col%) = AuxS
      Next
      AuxS = VetorB(k%)
      VetorB(k%) = VetorB(PosLinMaiorEle%)
      VetorB(PosLinMaiorEle%) = AuxS
    End If
    
    'pivo esta na coluna k?
    If k <> PosColMaiorEle% Then
      'nao, entao trocar coluna k com PosColMaiorEle
      For Lin% = 0 To NumEq - 1
        AuxS = MatrizA(Lin%, k%)
        MatrizA(Lin%, k%) = MatrizA(Lin%, PosColMaiorEle)
        MatrizA(Lin, PosColMaiorEle%) = AuxS
      Next
      AuxI% = VetorTrocaColX(k%)
      VetorTrocaColX%(k%) = VetorTrocaColX%(PosColMaiorEle)
      VetorTrocaColX%(PosColMaiorEle%) = AuxI%
    End If
    
    'pivotear
    For Lin% = k% + 1 To NumEq% - 1
      'elemento nulo?
      If MatrizA(Lin%, k%) <> 0 Then
        'nao
        'calcular os multiplicadores, guardar na matriz A parte inferior, sera a matriz L
        MatrizA(Lin%, k%) = MatrizA(Lin%, k%) / MatrizA(k%, k%)
        'atualizar a matriz A, parte superior, sera a matriz U, escalonada
        For Col = k + 1 To NumEq - 1
          MatrizA(Lin, Col) = MatrizA(Lin, Col) - MatrizA(Lin, k) * MatrizA(k, Col)
        Next
      End If
    Next

  Next 'fim da decomposicao LU - fim do for em k */


  'Sistema possivel?
  If TipoSist = SIST_POS_DET Then
    'sim
    'resolucao de Ly=b, y esta guardado no b
    For Col% = 0 To NumEq - 1
      For Lin% = 0 To Col - 1
        VetorB(Col) = VetorB(Col) - MatrizA(Col, Lin) * VetorB(Lin)
      Next
    Next
    'resolucao de Ux=y (y= b modificado), x esta guardado no b, x = b modificado pela 2a vez
    For Col% = NumEq - 1 To 0 Step -1
      For Lin = (Col + 1) To NumEq - 1
        VetorB(Col) = VetorB(Col) - MatrizA(Col, Lin) * VetorB(Lin)
      Next
      VetorB(Col) = VetorB(Col) / MatrizA(Col, Col)
    Next
    'atualizacao do vetor solucao b, atraves das trocas das colunas
    For Lin = 0 To NumEq - 2
      For Col = 0 To NumEq - 1
        'achei valor Lin dentro do vetor?
        If VetorTrocaColX(Col) = Lin Then
          'sim
          AuxI = VetorTrocaColX(Lin)
          VetorTrocaColX(Lin) = VetorTrocaColX(Col)
          VetorTrocaColX(Col) = AuxI
          AuxS = VetorB(Lin)
          VetorB(Lin) = VetorB(Col)
          VetorB(Col) = AuxS
          Exit For
        End If
      Next 'Col
    Next 'Lin
  Else
    'Sistema Possivel Indeterminado ou Impossivel
    Stop
  End If


End If

If TipoSol = 1 Then
'gauss parcial

  'Decomposicao LU
  For k% = 0 To NumEq% - 1
    'procurar maior elemento a partir da linha k e da coluna k
    AuxS! = Abs(MatrizA!(k%, k%))
    PosLinMaiorEle% = k%
    'PosColMaiorEle% = k%
    For Lin% = k% + 1 To NumEq% - 1
      'For Col% = k% To NumEq% - 1
        If (Abs(MatrizA!(Lin, k)) > Abs(AuxS!)) Then
          AuxS! = Abs(MatrizA!(Lin%, k))
          PosLinMaiorEle% = Lin%
          'PosColMaiorEle% = Col%
        End If
      'Next
    Next

    'pivo nulo?
    If (MatrizA!(PosLinMaiorEle%, k) = 0) Then
      'sim
      TipoSist = SIST_POS_IND_IMP
      Stop 'sistema possivel indeterminado
    End If

    'pivo esta na linha k?
    If (k% <> PosLinMaiorEle) Then
      'nao, entao trocar linha k com a linha iPosLinMaiorEle
      For Col% = 0 To NumEq - 1
        AuxS! = MatrizA(k%, Col%)
        MatrizA(k%, Col%) = MatrizA(PosLinMaiorEle%, Col%)
        MatrizA(PosLinMaiorEle%, Col%) = AuxS
      Next
      AuxS = VetorB(k%)
      VetorB(k%) = VetorB(PosLinMaiorEle%)
      VetorB(PosLinMaiorEle%) = AuxS
    End If
    
    'pivo esta na coluna k?
    'If k <> PosColMaiorEle% Then
    '  'nao, entao trocar coluna k com PosColMaiorEle
    '  For Lin% = 0 To NumEq - 1
    '    AuxS = MatrizA(Lin%, k%)
    '    MatrizA(Lin%, k%) = MatrizA(Lin%, PosColMaiorEle)
    '    MatrizA(Lin, PosColMaiorEle%) = AuxS
    '  Next
    '  AuxI% = VetorTrocaColX(k%)
    '  VetorTrocaColX%(k%) = VetorTrocaColX%(PosColMaiorEle)
    '  VetorTrocaColX%(PosColMaiorEle%) = AuxI%
    'End If
    
    'pivotear
    For Lin% = k% + 1 To NumEq% - 1
      'elemento nulo?
      If MatrizA(Lin%, k%) <> 0 Then
        'nao
        'calcular os multiplicadores, guardar na matriz A parte inferior, sera a matriz L
        MatrizA(Lin%, k%) = MatrizA(Lin%, k%) / MatrizA(k%, k%)
        'atualizar a matriz A, parte superior, sera a matriz U, escalonada
        For Col = k + 1 To NumEq - 1
          MatrizA(Lin, Col) = MatrizA(Lin, Col) - MatrizA(Lin, k) * MatrizA(k, Col)
        Next
      End If
    Next

  Next 'fim da decomposicao LU - fim do for em k */


  'Sistema possivel?
  If TipoSist = SIST_POS_DET Then
    'sim
    'resolucao de Ly=b
    For Col% = 0 To NumEq - 1
      For Lin% = 0 To Col - 1
        VetorB(Col) = VetorB(Col) - MatrizA(Col, Lin) * VetorB(Lin)
      Next
    Next
    'resolucao de Ux=y (y= b modificado), x esta guardado no b, x = b modificado pela 2a vez
    For Col% = NumEq - 1 To 0 Step -1
      For Lin = (Col + 1) To NumEq - 1
        VetorB(Col) = VetorB(Col) - MatrizA(Col, Lin) * VetorB(Lin)
      Next
      VetorB(Col) = VetorB(Col) / MatrizA(Col, Col)
    Next
    
    
    
    'atualizacao do vetor solucao b, atraves das trocas das colunas
    'For Lin = 0 To NumEq - 2
    '  For Col = 0 To NumEq - 1
    '    'achei valor Lin dentro do vetor?
    '    If VetorTrocaColX(Col) = Lin Then
    '      'sim
    '      AuxI = VetorTrocaColX(Lin)
    '      VetorTrocaColX(Lin) = VetorTrocaColX(Col)
    '      VetorTrocaColX(Col) = AuxI
    '      AuxS = VetorB(Lin)
    '      VetorB(Lin) = VetorB(Col)
    '      VetorB(Col) = AuxS
    '      Exit For
    '    End If
    '  Next 'Col
    'Next 'Lin


  Else
    'Sistema Possivel Indeterminado ou Impossivel
    Stop
  End If


End If


  
'inicio
  
If TipoSol = 2 Then
    
  'Decomposicao LU
  NumViaturas = iNumVertHiper
  'ReDim VDir%(NumEq - 1)
  
  For i% = 0 To 2 ^ (NumViaturas - 1) - 2
    VDir(i%) = 2 ^ (NumViaturas - 1) + i
  Next
  VDir(2 ^ (NumViaturas - 1) - 1) = 2 ^ (NumViaturas) - 1
  ct% = 0
  For i% = 2 ^ (NumViaturas - 1) To 2 ^ (NumViaturas) - 1
    VDir(i%) = ct%
    ct% = ct% + 1
  Next
  
 
  
  
  For k% = 0 To NumEq% - 1
    'primeiro elemento
    pivo = MatrizA(VDir(k%), k%)
    If Abs(pivo) < 0.000001 Then
      'pivo nulo
      For i% = k% + 1 To NumEq - 1
        If Abs(MatrizA(VDir(i%), k%)) > 0.000001 Then
          'trocar posicao de linha
          For j% = 0 To NumEq - 1
            aux! = MatrizA(VDir(k%), j%)
            MatrizA(VDir(k%), j%) = MatrizA(VDir(i%), j%)
            MatrizA(VDir(i%), j%) = aux!
          Next
          iaux% = VDir(k%)
          VDir(k%) = VDir(i%)
          VDir(i%) = iaux%
          Exit For
        End If
      Next
      If i% = NumEq Then
        'sistema impossivel
        Stop
      End If
    End If
    
    'sim, pivotear
    For i% = k% + 1 To NumEq% - 1
      If MatrizA(VDir(i%), k%) <> 0 Then
        fator = -MatrizA(VDir(i%), k%) / pivo
        For j% = k% To NumEq% - 1
          'pivoteamento
          If MatrizA(VDir(k%), j%) <> 0 Then
            MatrizA(VDir(i%), j%) = MatrizA(VDir(k%), j%) * fator + MatrizA(VDir(i%), j%)
            If Abs(MatrizA(VDir(i%), j%)) < 0.000001 Then
              MatrizA(VDir(i%), j%) = 0
            End If
          End If
        Next
        VetorB(VDir(i%)) = VetorB(VDir(k%)) * fator + VetorB(VDir(i%))


      End If
    Next
  Next 'fim da decomposicao LU - fim do for em k */

 
  
  
  'Sistema possivel?
  If TipoSist = SIST_POS_DET Then
    'sim
    For i% = NumEq - 1 To 0 Step -1
      For j% = i% + 1 To NumEq - 1
        VetorB(VDir%(i%)) = VetorB(VDir%(i%)) - (MatrizA(VDir%(i%), j%) * VetorB(VDir(j%)))
      Next
      VetorB(VDir(i%)) = VetorB(VDir(i%)) / MatrizA(VDir(i%), i%)
    Next
  Else
    'Sistema Possivel Indeterminado ou Impossivel
    Stop
  End If


End If
'fim
  
  
  

  resSistEq = TipoSist



End Function



Sub SetarValorVetorInt(cFin%, Vetor%(), Valor%)
  'estabelece um valor int para um vetor int
  Dim i%
  For i% = 0 To cFin%
    Vetor(i%) = Valor%
  Next
End Sub
 
Sub MultiplicarMatrizes(iLin1%, iCol1%, fMatriz1!(), iLin2%, iCol2%, fMatriz2!(), fMatriz3!())

'/* Multiplicar duas matrizes de qualquer tamanho */
'/* Parametros de entrada
' *   iLin1    = Numero de linhas da 1a matriz
' *   iCol1    = Numero de colunas da 1a matriz
' *   fMatriz1 = 1a matriz para multiplicacao
' *   iLin2    = Numero de linhas da 2a matriz
' *   iCol2    = Numero de colunas da 2a matriz
' *   fMatriz2 = 2a matriz para multiplicacao
' *   fMatriz3 = matriz resposta
' *
' * Retorno
' *   -1 = Erro na passagem dos parametros
' *    0 = ok
' */
'/* Multiplicar duas matrizes */

  Dim i%
  Dim j%
  Dim k%

  'da para fazer a multiplicacao?
  If iCol1 <> iLin2 Then
  'nao
    Stop
  End If
  For i = 0 To iLin1
    For j = 0 To iCol2
      For k = 0 To iCol1
        fMatriz3(i, j) = fMatriz3(i, j) + fMatriz1(i, k) * fMatriz2(k, j)
      Next
    Next
  Next

End Sub


Sub ImpTMS(titulo$, lIni, lFin, cIni, cFin, matriz!())
  Picture1.Scale (0, 0)-(60, 60)
  Picture1.Cls
  Picture1.Refresh
  mens$ = titulo + Chr$(13)
  For i% = lIni To lFin
    mens$ = mens$ + Format$(i%, "00")
    For j% = cIni To cFin
      Picture1.CurrentX = j% * 5
      Picture1.CurrentY = i% * 5
      Picture1.Print Format$(matriz(i%, j%), " #0.00000")
      mens$ = mens$ + Format$(matriz(i%, j%), " #0.00000")   'antes era chr$(10)+chr$(13)
    Next
    mens$ = mens$ + Chr$(13)
  Next
  MsgBox mens$
End Sub

Sub ImpTMI(titulo$, lIni, lFin, cIni, cFin, matriz%())
  mens$ = titulo + Chr$(13)
  For i% = lIni To lFin
    mens$ = mens$ + Format$(i%, "00")
    For j% = cIni To cFin
      mens$ = mens$ + Format$(matriz(i%, j%), "  000.00")   'antes era chr$(10)+chr$(13)
    Next
    mens$ = mens$ + Chr$(13)
  Next
  MsgBox mens$
End Sub

Sub ImpTVS(titulo$, lIni, lFin, Vetor!())
  mens$ = titulo + Chr$(13)
  For i% = lIni To lFin
    mens$ = mens$ + Format$(i%, "00") + Format$(Vetor!(i%), "  0.00000") + Chr$(13) 'antes era chr$(10)+chr$(13)
  Next
  MsgBox mens$
End Sub

Sub ImpTVI(titulo$, lIni, lFin, Vetor%())
  mens$ = titulo + Chr$(13)
  For i% = lIni To lFin
    mens$ = mens$ + Format$(i%, "00") + Format$(Vetor%(i%), "  0.00000") + Chr$(13) 'antes era chr$(10)+chr$(13)
  Next
  MsgBox mens$
End Sub

