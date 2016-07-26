@extends('layouts.movement')
@section('content')
<br/><br/>

<div class="row">
	<div class="col-sm-1">



</div>	

<div class="col-sm-12">

	
<p>Checkins and Checkouts</p>
<hr>

<br>
</div>	


<div class="col-sm-6 ">

<table class="table table-bodered table-responsive table-condensed">

                  <tr>
                    <td>Item</td><td>{{$item->tag_id}}</td>
                  </tr>
                  <tr>
                    <td>Description</td><td>{{$item->description}}</td>
                  </tr>
                  <tr>
                    <td>Category</td><td>{{$item->category}}</td>
                  </tr>

                  

                  @if($item->date_out != null)
                  <tr>
                    <td>Date Out</td><td>{{$item->date_out}}</td>
                  </tr>
                

                  <tr>
                    <td>Expected date back</td><td>{{$item->date_expected_back}}</td>
                  </tr>



                   <tr>
                    <td>Checked out by</td><td>{{$item->checked_out_by}}</td>
                  </tr>


                   <tr>
                    <td>Remarks</td><td>{{$item->remarks_out}}</td>
                  </tr>
                @endif
                
    
</table>

@if($item->date_out == null)
<a href="{{URL::to('movements/checkout/'.$item->id)}}" class="btn btn-primary btn-sm">Checkout</a>
@endif

@if($item->date_out != null)
<a href="{{URL::to('movements/checkin/'.$item->id)}}" class="btn btn-primary btn-sm">Checkin</a>  
@endif
</div>	



</div>





@stop