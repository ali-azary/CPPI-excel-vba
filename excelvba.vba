
Sub CPPI_mrkt_scenario()

Application.ScreenUpdating = False 'no updating the screen

Dim e As Integer 'Declare the variable e (interator) as an integer
Dim i As Integer 'Declare the variable i (interator) as an integer
Dim T As Integer 'Declare the variable T (# of periods) as an integer
Dim m As Integer 'Declare the variable m (CPPI multiplier) as an integer
Dim rf As Double 'Declare the variable rf (risk-free rate) as Double (for double-precision)
Dim mu As Double 'Declare the variable mu (average market return) as Double (for double-precision)
Dim sigma As Double 'Declare the variable sigma (market volatility) as Double (for double-precision)
Dim n As Integer 'Declare the variable n (number of iterations) as an integer
Dim mrkt As Double 'Declare the variable mrkt (market return index) as Double (for double-precision)
Dim CPPI As Double 'Declare the variable CPPI (CPPI index) as Double (for double-precision)
Dim rets() As Double 'array to store risky asset returns
Dim stats() 'market to store summary stats

'CLEAN THE DATA
Worksheets("Exercise_1_simulation").Range("B14:I1048576").ClearContents
Worksheets("Exercise_1_simulation").Range("C14:I1048576").NumberFormat = "0.00%"
Worksheets("Exercise_1_returns").Range("B3:AM1048576").ClearContents
Worksheets("Exercise_1_returns").Range("D3:AM1048576").NumberFormat = "0.00%"
Worksheets("Exercise_1_summary").Range("C3:D6").ClearContents
Worksheets("Exercise_1_summary").Range("E3:I9").ClearContents
Worksheets("Exercise_1_summary").Range("C3:D6").NumberFormat = "0.00%"
Worksheets("Exercise_1_summary").Range("E3:I9").NumberFormat = "0.00%"

Sheets("Exercise_1_simulation").Activate

'read input data
T = Range("C3").Value 'Store the defined number of periods in the variable T
m_i = Range("C10").Value 'initial m value
m_f = Range("C11").Value 'final m value
rf = Range("C4").Value 'Store the defined risk-free rate in the variable rf
mu = Range("C5").Value 'Store the defined market average return in the variable mu
sigma = Range("C6").Value 'Store the defined market volatility in the variable sigma
n = Range("C7").Value 'Store the defined number of simulations in the variable n

'write the headers
Cells(13, 3).Value = CStr(T) & "-period risk-free asset return"
Cells(13, 4).Value = CStr(T) & "-period risky asset return"
For m = m_i To m_f
    Cells(13, 4 + m).Value = CStr(T) & "-period CPPI return (m = " & CStr(m) & ")"
Next m
Range(Cells(13, 4 + m_f + 1), Cells(13, 16384)).ClearContents
Range(Cells(13, 2), Cells(13, 9)).WrapText = True
Range("B1:C1048576").Columns.AutoFit
Range(Cells(13, 4), Cells(13, 9)).Columns.AutoFit

Sheets("Exercise_1_summary").Activate

Cells(2, 3).Value = CStr(T) & "-period risk-free asset return"
Cells(2, 4).Value = CStr(T) & "-period risky asset return"
For m = m_i To m_f
    Cells(2, 4 + m).Value = CStr(T) & "-period CPPI return (m = " & CStr(m) & ")"
Next m
Range(Cells(2, 4 + m_f + 1), Cells(13, 16384)).ClearContents

Sheets("Exercise_1_simulation").Activate

'redimension risky asset return array to T periods
ReDim rets(T)

'risk-free asset return
rfret = (1 + rf) ^ T - 1

'redimension the stats arrayto hold summary statistics
ReDim stats(7, m_f)

'(FIRST) FOR LOOP TO PERFORM n NUMBER OF SIMULATIONS
For e = 1 To n 'Start a For loop from 1 to n (# of simulations)
    mrkt = 100 'Initial market net asset value (NAV) index
    Cells(13 + e, 2).Value = e 'write simulation no
    Sheets("Exercise_1_returns").Activate
    Cells(2 + e, 3).Value = e 'write simulation no
    
    'risky asset market returns for T periods
    For i = 1 To T
        ret = WorksheetFunction.Norm_Inv(Rnd(), mu, sigma)
        rets(i) = ret 'sotre random returns for use with CPPI later
        Cells(2 + e, 3 + i).Value = ret
        mrkt = mrkt * (1 + ret)
    Next i
    
    Sheets("Exercise_1_simulation").Activate
    Cells(13 + e, 3).Value = rfret 'write risk-free return
    Cells(13 + e, 4).Value = (mrkt / 100) - 1 'write risky return
    
    'loop over all m values for CPPI strategy
    For m = m_i To m_f
            
        CPPI = 100 'Initial CPPI strategy net asset value (NAV) index
    
        Floor = CPPI / (1 + rf) ^ T 'Calculate the initial CPPI floor for period i=0. This line follows a standard (and discrete) zero-coupon bond valuation with risk-free rate rf andnumber of periods T
    
        cushion = CPPI - Floor 'Calculate the initial cushion for period i=0
        
        'The command WorksheetFunction allows you to use functions from the Excel sheet in VBA. In this case, the output of the Excel function MAX() is the maximum number in its arguments
        invest_mrkt = WorksheetFunction.Max(cushion * m, 0) 'Calculate the initial portfolio proportion invested in the risk asset (with market returns) in period i=0. When the cushion is negative, all the portfolio is invested in the risk-free asset
    
        invest_rf = CPPI - invest_mrkt 'Calculate the initial portfolio proportion invested in the risk-free asset in period i=0
    
        '(SECOND) FOR LOOP TO PERFORM NAV INDEX CALCULATIONS IN T NUMBER OF INVESTMENT PERIODS. THIS LOOP IS WITHIN THE FIRST LOOP.
        For i = 1 To T 'Start a For loop from 1 to T (# of investment periods)
    
            'GENERATE THE MARKET RETURN FOR THIS INVESTMENT PERIOD AND CALCULATE THE INDEX FOR MARKET AND CPPI INDEXES IN INVESTMENT PERIOD i
            'The command WorksheetFunction allows you to use functions from the Excel sheet in VBA. The worksheet function RAND() generates a random number between 0 and 1.
            ret = rets(i) 'random returns from risky asset
            
            CPPI = invest_rf * (1 + rf) + invest_mrkt * (1 + ret) 'Calculate the variation in the CPPI index based on the normally distributed random market return of this period (ret)
            
              
            'CALCULATE THE RESULTING FLOOR, CUSHION, MARKET NAV INDEX AND CPPI NAV INDEX FOR THE INVESTMENT PERIOD i
            Floor = 100 / (1 + rf) ^ (T - i) 'Calculate the floor in period i. This line follows a standard (and discrete) zero-coupon bond valuation with risk-free rate rf andnumber of periods T - i
    
            cushion = CPPI - Floor 'Calculate the  cushion for period i
            
            'The command WorksheetFunction allows you to use functions from the Excel sheet in VBA.
            invest_mrkt = WorksheetFunction.Max(cushion * m, 0) 'Calculate the portfolio proportion invested in the risk asset (with market returns) in period i. When the cushion is negative, all the portfolio is invested in the risk-free asset
            invest_mrkt = WorksheetFunction.Min(invest_mrkt, CPPI) 'Prevent investing more than the total CPPI NAV (no leverage) in the risk asset in case the cushion*m is larger than the CPPI NAV in investment period i
    
            invest_rf = CPPI - invest_mrkt 'Calculate the portfolio proportion invested in the risk-free asset in period i
    
        Next i 'The next investment cycle starts with the next i (current i +1)
        
        
        Cells(13 + e, 4 + m).Value = (CPPI / 100) - 1 'Calculate T-period return for the CPPI NAV index in simulation e and paste the result in the corresponding row (12 + 3) and column (D)
    stats(7, m) = stats(7, m) + IIf(Cells(13 + e, 4 + m).Value < Cells(13 + e, 4).Value, 1, 0) / n 'probability of CPPI return less than risky asset return
    Next m
Next e
    
' summary stats
For m = m_i To m_f
    stats(1, m) = WorksheetFunction.Average(Columns(4 + m)) 'average return for CPPI for each m value
    stats(2, m) = WorksheetFunction.StDev(Columns(4 + m)) 'standard deviations
    stats(3, m) = WorksheetFunction.Min(Columns(4 + m)) 'min return
    stats(4, m) = WorksheetFunction.Max(Columns(4 + m)) 'max return
    stats(5, m) = WorksheetFunction.CountIf(Columns(4 + m), "<0") / n 'probability of negative CPPI return
    stats(6, m) = WorksheetFunction.CountIf(Columns(4 + m), "<" & rfret) / n 'probability of CPPI return less than risk-free
Next m

'for risky asset
mean_risky = WorksheetFunction.Average(Range("D14:D1048576"))
std_risky = WorksheetFunction.StDev(Range("D14:D1048576"))
min_risky = WorksheetFunction.Min(Range("D14:D1048576"))
max_risky = WorksheetFunction.Max(Range("D14:D1048576"))

'write summary stats
Sheets("Exercise_1_summary").Activate

Cells(3, 4).Value = mean_risky
Cells(4, 4).Value = std_risky
Cells(5, 4).Value = min_risky
Cells(6, 4).Value = max_risky
Range(Cells(3, 3), Cells(6, 3)).Value = rfret

For m = m_i To m_f
    For i = 1 To 7
        Cells(2 + i, 4 + m).Value = stats(i, m)
    Next i
Next m
    

End Sub

