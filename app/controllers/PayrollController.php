<?php

class PayrollController extends \BaseController {

    /**
     * Display a listing of branches
     *
     * @return Response
     */
    public function index()
    {
        $accounts = Account::all();

        return View::make('payroll.index', compact('accounts'));
    }

    public function preview_payroll()
    {

        $employees = DB::table('employee')
                  ->where('in_employment','=','Y')
                  ->get();

        //print_r($accounts);

        Audit::logaudit('Payroll', 'preview', 'previewed payroll');


        return View::make('payroll.preview', compact('employees'));
    }

    public function valid()
    {
        $period = Input::get('period');

        //print_r($accounts);

        return View::make('payroll.valid', compact('period'));
    }

    /**
     * Show the form for creating a new branch
     *
     * @return Response
     */
    public function create()
    {
        $employees = DB::table('employee')
                  ->where('in_employment','=','Y')
                  ->get();
        $period = Input::get('period');
        $account = Input::get('account');

        //print_r($accounts);

        Audit::logaudit('Payroll', 'preview', 'previewed payroll');

        return View::make('payroll.preview', compact('employees','period','account'));
    }

    public function del_exist()
    {
    $postedit = Input::all();
    $part1    = $postedit['period1'];
    $part2    = $postedit['period2'];
    $part3    = $postedit['period3'];

    $period   = $part1.$part2.$part3;  

     DB::table('employee_deductions')
              ->join('transact_deductions','employee_deductions.id','=','transact_deductions.employee_deduction_id')
              ->where('financial_month_year', '=', $period)
              ->where(function($query){
                $query->where('formular','=','One Time')
                      ->orWhere('formular','=','Instalments');
               })
              ->increment('instalments');

    DB::table('earnings')
              ->join('transact_earnings','earnings.id','=','transact_earnings.earning_id')
              ->where('financial_month_year', '=', $period)
              ->where(function($query){
                $query->where('formular','=','One Time')
                      ->orWhere('formular','=','Instalments');
               })
              ->increment('instalments');
    
    $data     = DB::table('transact')->where('financial_month_year',$period)->delete(); 
    $data2    = DB::table('transact_allowances')->where('financial_month_year', '=', $period)->delete();
    $data3    = DB::table('transact_deductions')->where('financial_month_year', '=', $period)->delete();
    $data4    = DB::table('transact_earnings')->where('financial_month_year', '=', $period)->delete();
    $data5    = DB::table('transact_overtimes')->where('financial_month_year', '=', $period)->delete();
    $data6    = DB::table('transact_reliefs')->where('financial_month_year', '=', $period)->delete();

   
    if($data > 0){
      return 0;
    }else{
      return 1;
    }
    

    exit();
    }

    public function display(){
      $display = "";
      $postedit = Input::all();
      $part1    = $postedit['period1'];
      $part2    = $postedit['period2'];
      $part3    = $postedit['period3'];

      $fperiod   = $part1.$part2.$part3; 
      $employees = DB::table('employee')
                  ->where('in_employment','=','Y')
                  ->get();
        
        $i=1;
        foreach($employees as $employee){
        $salary = number_format($employee->basic_pay,2);
        $benefits = number_format(Payroll::total_benefits($employee->id,$fperiod),2);
        $gross = number_format(Payroll::gross($employee->id,$fperiod),2);
        $paye = number_format(Payroll::tax($employee->id,$fperiod),2);
        $nssf = number_format(Payroll::nssf($employee->id,$fperiod),2);
        $nhif = number_format(Payroll::nhif($employee->id,$fperiod),2);
        $deductions = number_format(Payroll::deductions($employee->id,$fperiod),2);
        $total_deductions = number_format(Payroll::total_deductions($employee->id,$fperiod),2);
        $net = number_format(Payroll::net($employee->id,$fperiod),2);

        $display .="
        <tr>

          <td> $i </td>
          <td >$employee->personal_file_number</td>
          <td>$employee->first_name $employee->last_name </td>
          <td align='right'>$salary</td>
          <td align='right'>$benefits</td>
          <td align='right'>$gross</td>
          <td align='right'>$paye</td>
          <td align='right'>$nssf</td>
          <td align='right'>$nhif</td>
          <td align='right'>$deductions</td>
          <td align='right'>$total_deductions</td>
          <td align='right'>$net</td>
          
        </tr>
        ";
         $i++;
         
        } 
        return $display;
        exit();

    }

    /**
     * Store a newly created branch in storage.
     *
     * @return Response
     */
    public function store()
    {
        $employees = DB::table('employee')
                  ->where('in_employment','=','Y')
                  ->get();

        foreach ($employees as $employee) {
        $payroll = new Payroll;

        $payroll->employee_id = $employee->personal_file_number;
        $payroll->basic_pay = $employee->basic_pay; 
        $payroll->earning_amount = Payroll::total_benefits($employee->id,Input::get('period'));
        $payroll->taxable_income = Payroll::gross($employee->id,Input::get('period'));
        $payroll->paye = Payroll::tax($employee->id,Input::get('period'));
        $payroll->nssf_amount = Payroll::nssf($employee->id,Input::get('period'));
        $payroll->nhif_amount = Payroll::nhif($employee->id,Input::get('period'));
        $payroll->other_deductions = Payroll::deductions($employee->id,Input::get('period'));
        $payroll->total_deductions = Payroll::total_deductions($employee->id,Input::get('period'));
        $payroll->net = Payroll::net($employee->id,Input::get('period'));
        $payroll->financial_month_year = Input::get('period');
        $payroll->account_id = Input::get('account');
        $payroll->save();
        }


    
        $allws = DB::table('employee_allowances')
            ->join('allowances', 'employee_allowances.allowance_id', '=', 'allowances.id')
            ->join('employee', 'employee_allowances.employee_id', '=', 'employee.id')
            ->select('employee.id as eid','employee_allowances.id as id','allowance_name','allowance_id','allowance_amount')
            ->get();

        foreach($allws as $allw){
        DB::table('transact_allowances')->insert(
        ['employee_id' => $allw->eid, 
        'employee_allowance_id' => $allw->id, 
        'allowance_name' => $allw->allowance_name,
        'allowance_id' => $allw->allowance_id,
        'allowance_amount' => $allw->allowance_amount,
        'financial_month_year'=>Input::get('period'),
        ]
        );
        }

        $period = Input::get('period');
        $part = explode("-", $period);
        $start = $part[1]."-".$part[0]."-01";
        $end  = date('Y-m-t', strtotime($start));

        $deds = DB::table('employee_deductions')
            ->join('deductions', 'employee_deductions.deduction_id', '=', 'deductions.id')
            ->join('employee', 'employee_deductions.employee_id', '=', 'employee.id')
            ->where('instalments','>',0)
            ->where('first_day_month','<=',$start)
            ->where('last_day_month','>=',$start)
            ->select('employee.id as eid','employee_deductions.id as id','deduction_name','deduction_id','formular','instalments','deduction_amount')
            ->get();

        $count = DB::table('employee_deductions')
            ->join('deductions', 'employee_deductions.deduction_id', '=', 'deductions.id')
            ->join('employee', 'employee_deductions.employee_id', '=', 'employee.id')
            ->where('instalments','>',0)
            ->where('first_day_month','<=',$start)
            ->where('last_day_month','>=',$start)
            ->select('employee.id as eid','employee_deductions.id as id','deduction_name','deduction_id','formular','instalments','deduction_amount')
            ->count();
        
        if($count>0){
        foreach($deds as $ded){
        DB::table('transact_deductions')->insert(
        ['employee_id' => $ded->eid, 
        'employee_deduction_id' => $ded->id, 
        'deduction_name' => $ded->deduction_name,
        'deduction_id' => $ded->deduction_id,
        'deduction_amount' => $ded->deduction_amount,
        'financial_month_year'=>Input::get('period'),
        ]
        );
        }

        DB::table('employee_deductions')
             ->where(function($query){
                $query->where('formular','=','One Time')
                      ->orWhere('formular','=','Instalments');
               })
             ->where('instalments','>',0)
             ->decrement('instalments');
        
        }


        $earns = DB::table('earnings')
            ->join('employee', 'earnings.employee_id', '=', 'employee.id')
            ->where('instalments','>',0)
            ->where('first_day_month','<=',$start)
            ->where('last_day_month','>=',$start)
            ->select('earnings.employee_id','earnings.id as id','earnings_name','earnings_amount','formular','instalments')
            ->get();

        $ct = DB::table('earnings')
            ->join('employee', 'earnings.employee_id', '=', 'employee.id')
            ->where('instalments','>',0)
            ->where('first_day_month','<=',$start)
            ->where('last_day_month','>=',$start)
            ->select('earnings.employee_id','earnings.id as id','earnings_name','earnings_amount','formular','instalments')
            ->count();

        if($ct>0){
        foreach($earns as $earn){
        DB::table('transact_earnings')->insert(
        ['employee_id' => $earn->employee_id, 
        'earning_id' => $earn->id,
        'earning_name' => $earn->earnings_name,
        'earning_amount' => $earn->earnings_amount,
        'financial_month_year'=>Input::get('period'),
        ]
        );
        }

        DB::table('earnings')
             ->where(function($query){
                $query->where('formular','=','One Time')
                      ->orWhere('formular','=','Instalments');
               })
             ->where('instalments','>',0)
             ->decrement('instalments');

        }

        $overtimes = DB::table('overtimes')
            ->join('employee', 'overtimes.employee_id', '=', 'employee.id')
            ->select('overtimes.employee_id','overtimes.type','overtimes.period','overtimes.amount')
            ->get();

        foreach($overtimes as $overtime){
        DB::table('transact_overtimes')->insert(
        ['employee_id' => $overtime->employee_id, 
        'overtime_type' => $overtime->type,
        'overtime_period' => $overtime->period,
        'overtime_amount' => $overtime->amount,
        'financial_month_year'=>Input::get('period'),
        ]
        );
        }

        $rels = DB::table('employee_relief')
            ->join('relief', 'employee_relief.relief_id', '=', 'relief.id')
            ->join('employee', 'employee_relief.employee_id', '=', 'employee.id')
            ->select('employee.id as eid','employee_relief.id as id','relief_name','relief_id','relief_amount')
            ->get();

        foreach($rels as $rel){
        DB::table('transact_reliefs')->insert(
        ['employee_id' => $rel->eid, 
        'employee_relief_id' => $rel->id, 
        'relief_name' => $rel->relief_name,
        'relief_id' => $rel->relief_id,
        'relief_amount' => $rel->relief_amount,
        'financial_month_year'=>Input::get('period'),
        ]
        );
        }

        $period = Input::get('period'); 
        Audit::logaudit('Payroll', 'process', 'processed payroll for '.$period);
    
    return Redirect::route('payroll.index')->withFlashMessage('Payroll successfully processed!');
         

    }

    

    /**
     * Display the specified branch.
     *
     * @param  int  $id
     * @return Response
     */
    public function show($id)
    {
        $payroll = Payroll::findOrFail($id);

        return View::make('payroll.show', compact('payroll'));
    }

    /**
     * Show the form for editing the specified branch.
     *
     * @param  int  $id
     * @return Response
     */
    public function edit($id)
    {
        $deduction = Deduction::find($id);

        return View::make('deductions.edit', compact('deduction'));
    }

    /**
     * Update the specified branch in storage.
     *
     * @param  int  $id
     * @return Response
     */
    public function update($id)
    {
        $deduction = Deduction::findOrFail($id);

        $validator = Validator::make($data = Input::all(), Deduction::$rules, Deduction::$messages);

        if ($validator->fails())
        {
            return Redirect::back()->withErrors($validator)->withInput();
        }

        $deduction->deduction_name = Input::get('name');
        $deduction->update();

        return Redirect::route('deductions.index');
    }

    /**
     * Remove the specified branch from storage.
     *
     * @param  int  $id
     * @return Response
     */
    public function destroy($id)
    {
        Deduction::destroy($id);

        return Redirect::route('deductions.index');
    }

}
