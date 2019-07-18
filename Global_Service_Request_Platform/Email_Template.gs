<!DOCTYPE html>
<html>
  <head>
    <style>
      table, th, td {
        border: 1px solid black;
      }
      table tr:last-child
      {
        display:none !important;
      }
	  h4{
		border-bottom:2px solid #ff9400;
	  }
    </style>
    <base target="_top">
  </head>
  <body>
    <h4>
    <p>Hello,</p>

       </h4>
        <p>Have a great day.</p>
        <p>We acknowledge your concern.</p>
        <p>Your request has been auto resolved successfully. <?= data[2] ?></p>
      <table>
        <tr>
          <th bgcolor="#FF9400">S/N</th>
          <th bgcolor="#FF9400">Key</th>
          <th bgcolor="#FF9400">Value</th>
        </tr>
        <? for (var i = 0; i < data[0].length; i++) { ?>
          <tr>
              <td><?=i+1 ?></td>
              <td><?= data[0][i] ?></td>
              <td><?= data[1][i] ?></td>
          </tr>
        <? } ?>
         </table>
      <hr/>
      <p><strong style="color:#343434;"><p>Best Regards,</p>RE<span style="color:#ff9400;border-top:2px solid #ff9400;">A</span>N Cloud IT Operations</strong></p>
      <p>[NOTE]: This email is auto generated. For any further clarification or queries please feel free to reply this email.</p>
    </body>
</html>
