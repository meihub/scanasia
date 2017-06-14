<?php 
//Setup connection to MySql
$servername = "localhost";
$username = "root";
$password = "daniel1030";
$dbname = "scanasia";

// Create connection
$conn = new mysqli($servername, $username, $password, $dbname);
// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 

// Select all columns from Product table
$sql = "SELECT * FROM Product";
$result = $conn->query($sql);

?>


<html lang="en">

<head>
  <meta charset="utf-8">

  <title>
    <?php echo $product_name?>
  </title>
  <meta name="description" content="The HTML5 Herald">
  <meta name="author" content="SitePoint">

  <link rel="stylesheet" href="css/styles.css?v=1.0">
</head>
<body>
  <!Place Scan Asia Logo HERE> 
  <div style="background:green;height:100px">
  
  </div>
  <!create a div box for Headers> 
  <div style="background:orange;height:100px">
      <!First Header> 
      <div style="height:50%;border-style:solid;border-color:black" >
      1st Header box goes HERE
     </div>


      <!2nd Header> 
      <div style="height:35%;border-style:solid;border-color:red">
      2nd Header box goes HERE
     </div>

  </div>
   <!For TESTING purposes only> 
  Total number of rows:
  <?php echo $result->num_rows?>
  
  <div style="background-color:grey;color:black;height:500px;width:100%;position:relative;">
  
        <div style="background-color:yellow;padding:20px 20px 20px 20px;position:absolute;width:230px;top:0px;bottom:0px;left:0px;height:auto">SCREWS
            <!create a list of sub 2 product categories HERE> 
        </div>
        

        <div style="background-color:pink;padding:20px 20px 20px 20px;position:absolute;top:0px;bottom:0px;left:270px;right:0;height:auto">SELF TAPPING SCREWS
            <!create div boxes side by side as image> 

             <!then populate the description for each image>   
                  <table>
                    <thead>
                      <th>Product Name</th>
                      <th>Product Type</th>
                    </thead>
                    <tbody>
                      <?php 
                        while($row = $result->fetch_assoc()) {?>
                      <tr>
                        <td><?php echo $row["product name"] ?></td>
                        <td><?php echo $row["product type"] ?></td>
                      </tr>
                      <?php } ?>
                    </tbody>
                  </table>
              </div>
  
  </div>
 
  <script src="js/scripts.js"></script>
</body>

</html>

<?php

$conn->close();

?>