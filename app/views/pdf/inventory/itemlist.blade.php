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

<div class="footer">
     <p class="page">Page <?php $PAGE_NUM ?></p>
   </div>


	<div class="content" style='margin-top:0px;'>
@if($item == 'all')
<div align="center"><strong>Item Lists Report </strong></div><br>
<table class="table table-bordered" border='1' cellspacing='0' cellpadding='0'>

      <tr>
        
        <td width='20'><strong># </strong></td>
        <td><strong>Name </strong></td>
        <td><strong>Description </strong></td>
        <td><strong>Category</strong></td>
        <td><strong>Store Keeping Unit</strong></td>
        <td><strong>Tag ID</strong></td>
      </tr>
      <?php $i =1; ?>
      @foreach($items as $item)
      <tr>
       <td td width='20'>{{$i}}</td>
       <td>{{ $item->name }}</td>
       <td>{{ $item->description }}</td>  
       <td>{{ $item->category }}</td>
       <td>{{ $item->sku }}</td> 
       <td>{{ $item->tag_id }}</td> 
      </tr>
      <?php $i++; ?>
   
    @endforeach

     
    </table>

<br><br>
@else
<div align="center"><strong>Report for {{$itm->name}} </strong></div><br>
<table class="table table-bordered" border='1' cellspacing='0' cellpadding='0'>

      <tr>
        <td><strong>Name </strong></td>
        <td><strong>Description </strong></td>
        <td><strong>Category</strong></td>
        <td><strong>Store Keeping Unit</strong></td>
        <td><strong>Tag ID</strong></td>

      </tr>
      <tr>
       <td>{{ $itm->name }}</td>
       <td>{{ $itm->description }}</td> 
       <td>{{ $itm->category }}</td>
       <td>{{ $itm->sku }}</td> 
       <td>{{ $itm->tag_id }}</td> 
      </tr>

     
    </table>

<br><br>
@endif
  
</div>


</body>

</html>



