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
    <h4>Hi Team,<br><br> 
            <?= data[1] ?>
                    <p>
        </p></h4>
        <p>
        </p>
      <table>
        <tr>
          <th>S.N.</th>
          <th>Member Email</th>
        </tr>
        <? for (var i = 0; i < data[0].length; i++) { ?>
          <tr>
              <td><?=i+1 ?></td>
              <td><?= data[0][i][0] ?></td>
          </tr>
        <? } ?>
      </table>
      <hr/>
      <p><strong style="color:#343434;"><p>Best Regards,</p>RE<span style="color:#ff9400;border-top:2px solid #ff9400;">A</span>N Cloud IT Operations</strong></p>
      <p>[NOTE]: The Report is now automatically generated.</p>
    </body>
</html>

