@extends('layouts.erp')
@section('content')
<br><br/>
<div class="row">
                      

    <div class="col-lg-12">

      <h4>Inventory Reports</h4>

    </div>                 

    </div>



<hr>
<div class="row">
              						
<div class="col-lg-2"></div>
	<div class="col-lg-5">

@if (Session::get('notice'))
            <div class="alert alert-info">{{ Session::get('notice') }}</div>
        @endif
		
    <form target="blank" method="POST" action="{{{ URL::to('invreports') }}}" accept-charset="UTF-8">

           <div class="form-group">
            <label for="username">Report <span style="color:red">*</span> :</label>
            <select name="report_type" class="form-control" required id="report_type">
               <option value=""></option>
               <option value="checkout">Checkout report</option>
               <option value="checkin">Checkin report</option>
               <option value="booking">Bookings report</option>
               <option value="maintenance">Maintenance report</option>
               <option value="item_list">Item List report</option>
               <option value="store_list">Store report</option>
            </select>
        </div>

        <div class="form-group" id="clients">
            <label for="username">Clients <span style="color:red">*</span> :</label>
            <select name="client" class="form-control">
               <option value=""></option>
               @foreach($clients as $client)
               <option value="{{$client->id}}">{{$client->name}}</option>
               @endforeach
              
            </select>
        </div>

        <div class="form-group" id="booked_events">
            <label for="username">Event <span style="color:red">*</span> :</label>
            <select name="event" class="form-control">
               <option value=""></option>
               @foreach($bookings as $booking)
               <option value="{{$booking->event}}">{{$booking->event}}</option>
               @endforeach
              
            </select>
        </div>

         <div class="form-group" id="item" class="item">
            <label for="username">Item <span style="color:red">*</span> :</label>
            <select name="items" class="form-control">
               <option value="all">All</option>
               @foreach($items as $item)
               <option value="{{$item->id}}">{{$item->name}}</option>
               @endforeach
              
            </select>
        </div>



         <div class="form-group" id="store" class="store">
            <label for="username">Store <span style="color:red">*</span> :</label>
            <select name="store" class="form-control">
               <option value="all">All</option>
               @foreach($stores as $store)
               <option value="{{$store->id}}">{{$store->name}}</option>
               @endforeach
              
            </select>
        </div>

        <div class="form-group" id="maintenance_test">
            <label for="username">Maintenance Test <span style="color:red">*</span> :</label>
            <select name="tests" class="form-control">
               <option value="all">All</option>
               @foreach($tests as $test)
               <option value="{{$test->id}}">{{$test->name}}</option>
               @endforeach
            </select>
        </div>


         <div class="form-group" id="start_date">
                        <label for="username">Start Date</label>
                        <div class="right-inner-addon ">
                        <i class="glyphicon glyphicon-calendar"></i>
                        <input class="form-control datepicker21"  readonly="readonly" placeholder="" type="text" name="start_date"  >
                        </div>
          </div>


           <div class="form-group" id="end_date">
                        <label for="username">End Date</label>
                        <div class="right-inner-addon ">
                        <i class="glyphicon glyphicon-calendar"></i>
                        <input class="form-control datepicker21"  readonly="readonly" placeholder="" type="text" name="end_date"  >
                        </div>
          </div>


         

        <div class="form-actions form-group">
        
          <button type="submit" class="btn btn-primary btn-sm">View Report</button>
        </div>


    </form>
		
    
	</div>


</div>


<script type="text/javascript">


  $(document).ready(function(){

    $('#store').hide();
    $('#maintenance_test').hide();
    $('#booked_events').hide();
    $('#item').hide();
    $('#clients').hide();
    $('#start_date').hide();
    $('#end_date').hide();


    $('#report_type').change(function(){

        if($("#report_type").val() == 'item_list'){
           $('#store').hide();
           $('#booked_events').hide();
           $('#maintenance_test').hide();
           $('#clients').hide();
           $('#item').show();
           $('#start_date').hide();
           $('#end_date').hide();
        }

        else if($("#report_type").val() == ''){
           $('#store').hide();
           $('#booked_events').hide();
           $('#maintenance_test').hide();
           $('#clients').hide();
           $('#item').hide();
           $('#start_date').hide();
           $('#end_date').hide();
        }

        else if($("#report_type").val() == 'store_list'){
           $('#item').hide();
           $('#store').show();
           $('#maintenance_test').hide();
           $('#booked_events').hide();
           $('#clients').hide();
           $('#start_date').hide();
           $('#end_date').hide();
           $('#booked_items').hide();
        }

        else if($("#report_type").val() == 'booking'){
          $('#store').hide();
          $('#item').hide();
          $('#maintenance_test').hide();
          $('#booked_events').show();
          $('#clients').show();
          $('#start_date').hide();
          $('#end_date').hide();
        }

        else if($("#report_type").val() == 'maintenance'){
          $('#store').hide();
          $('#item').hide();
          $('#maintenance_test').show();
          $('#booked_events').hide();
          $('#clients').hide();
          $('#start_date').show();
          $('#end_date').show();
        }

        else if($("#report_type").val() == 'checkout'){
          $('#store').hide();
          $('#item').hide();
          $('#maintenance_test').hide();
          $('#booked_events').hide();
          $('#clients').hide();
          $('#start_date').show();
          $('#end_date').show();
        }
        
        else if($("#report_type").val() == 'checkin'){
          $('#store').hide();
          $('#item').hide();
          $('#maintenance_test').hide();
          $('#booked_events').hide();
          $('#clients').hide();
          $('#start_date').show();
          $('#end_date').show();
        }

    });


  });

</script>

@stop