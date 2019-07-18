<!DOCTYPE html>
<html>
  <head>
    <style>
      table, th, td {
        border: 1px solid black;
      }
	  h4{
		border-bottom:2px solid #ff9400;
	  }
    </style>
    <base target="_top">
  </head>
  <body>
    <h4>Hi <?= data[2] ?>,<br><br> 
        </h4>
        <p>Have a great day!</p>
        <p><?= data[0] ?></p>
        <p></p>
        <p></p>
       <table>
        <tr>
          <th bgcolor="#CDD0CC">S.N.</th>
          <th bgcolor="#CDD0CC">Comments</th>
        </tr>
          <tr>
              <td>1</td>
              <td><?= data[5] ?></td>
          </tr>
      </table> 
        <p></p>
        <p></p>        
      <table>
        <tr>
          <th bgcolor="#3498DB">S.N.</th>
          <th bgcolor="#3498DB">Menu</th>
          <th bgcolor="#3498DB">Details</th>
        </tr>
        <? for (var i = 0; i < data[3].length; i++) { ?>
          <tr>
              <td><?=i+1 ?></td>
              <td><?= data[3][i][0] ?></td>
              <td><?= data[4][i][0] ?></td>
          </tr>
        <? } ?>
      </table>
      <p></p>
      <p></p>
       <table>
        <tr>
          <th bgcolor="#E67E22">S.N.</th>
          <th bgcolor="#E67E22">Jita Ticket ID</th>
          <th bgcolor="#E67E22">Status</th>
        </tr>
          <tr>
              <td>1</td>
              <td><?= data[1] ?></td>
              <td bgcolor="#27AE60">Done</td>
          </tr>
      </table>     
      <hr/>
      <p><strong style="color:#343434;"><p>Best Regards,</p>RE<span style="color:#ff9400;border-top:2px solid #ff9400;">A</span>N Cloud IT Operations</strong></p>
      <p>[NOTE]: The Report is now automatically generated.</p>
    </body>
</html>

