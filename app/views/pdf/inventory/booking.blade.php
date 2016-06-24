<?php


function asMoney($value) {
  return number_format($value, 2);
}

?>
<html >



<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

<style type="text/css">

table {
  max-width: 100%;
  background-color: transparent;
}
th {
  text-align: left;
}
.table {
  width: 100%;
  margin-bottom: 2px;
}
hr {
  margin-top: 1px;
  margin-bottom: 2px;
  border: 0;
  border-top: 2px dotted #eee;
}

body {
  font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
  font-size: 12px;
  line-height: 1.428571429;
  color: #333;
  background-color: #fff;
}



 @page { margin: 170px 30px; }
 .header { position: fixed; left: 0px; top: -150px; right: 0px; height: 150px;  text-align: center; }
 .content {margin-top: -100px; margin-bottom: -150px}
 .footer { position: fixed; left: 0px; bottom: -180px; right: 0px; height: 50px;  }
 .footer .page:after { content: counter(page, upper-roman); }



</style>

</head>

<body>

  <div class="header">
     <table >

      <tr>


       
        <td style="width:150px">

            <img src="{{public_path().'/uploads/logos/'.$organization->logo}}" alt="logo" width="80%">

    
        </td>

        <td>
        <strong>
          {{ strtoupper($organization->name)}}<br>
          </strong>
          {{ $organization->phone}} |
          {{ $organization->email}} |
          {{ $organization->website}}<br>
          {{ $organization->address}}
       

        </td>
        

      </tr>


      <tr>

        <hr>
      </tr>



    </table>
   </div>

<br>

<?php
  $d=strtotime($from);
  $d1=strtotime($to);
?>

<div class="footer">
     <p class="page">Page <?php $PAGE_NUM ?></p>
   </div>

	<div class="content" style='margin-top:0px;'>

   <div align="center"><strong>Booking Report for {{$booking->client->name}} </strong></div><br>
    <table class="table table-bordered" border='1' cellspacing='0' cellpadding='0'>
            <tr>
              <td>Event</td><td>{{$booking->event}}</td>
            </tr>
            <tr>
              <td>Start Date</td><td>{{$booking->start_date}}</td>
            </tr>
            <tr>
              <td>End Date</td><td>{{$booking->end_date}}</td>
            </tr>
            <tr>
              <td>Client</td><td>{{$booking->client->name}}</td>
            </tr>
             <tr>
              <td>Venue</td><td>{{$booking->venue}}</td>
            </tr>
             <tr>
              <td>Tech Lead</td><td>{{$booking->lead}}</td>
            </tr>
          </table>

<br><br>
<div align="center"><strong>Booked Items </strong></div><br>
    
<table border='1' cellspacing='0' cellpadding='0' class="table table-condensed table-bordered">
            
            <tr style="font-weight:bold">
              <td>Item</td><td>Tag </td><td>SKU</td>
            </tr>
            @foreach(Booking::getItems($booking->id) as $item)
            <tr>
              <td>{{Item::getItemName($item->item_id)}}</td>
              <td>{{Item::getItemTag($item->item_id)}}</td>
              <td>{{Item::getItemSku($item->item_id)}}</td>
            </tr>
            @endforeach
          </table>




   
</div>


</body>

</html>



