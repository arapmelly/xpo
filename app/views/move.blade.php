@extends('layouts.movement')
@section('content')
<br/><br/>

<div class="row">
	<div class="col-sm-1">



</div>	

<div class="col-sm-12">

	
<p>Checkout</p>
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

                  

 
</table>


<form method="POST" action="{{{ URL::to('checks') }}}" accept-charset="UTF-8">
   
    <fieldset>
        

         

         <input type="hidden" name="item_id" value="{{$item->id}}"/>

        <div class="form-group">
            <label for="username">Date Out <span style="color:red">*</span> :</label>
            <input class="form-control datepicker21" placeholder="" type="text" name="date_out" id="date_out" value="{{date('Y-m-d')}}" required>
        </div>


          <div class="form-group">
            <label for="username">Date Expected Back :</label>
            <input class="form-control datepicker21" placeholder="" type="text" name="expected_date_back" id="expected_date_back" value="{{{ Input::old('expected_date_back') }}}" required>
        </div>


         <div class="form-group">
            <label for="username">Remarks:</label>
          <textarea name="remarks_out" class="form-control"></textarea>
        </div>


        

       
        <div class="form-actions form-group">
        
          <button type="submit" class="btn btn-primary btn-sm"> Checkout</button>
        </div>

    </fieldset>
</form>



</div>	



</div>





@stop