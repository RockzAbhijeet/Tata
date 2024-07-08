<?php
include 'connect.php';
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tata Motors Tools and Machinery</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <nav class="navbar">
        <div class="container">
            <div class="logo">
                <img src="./logo.jpg" alt="Tata Motors Logo">
            </div>
            <ul class="nav-links">
                <li><a href="#home">Home</a></li>
                <li><a href="#vehicles">Vehicles</a></li>
                <li><a href="#services">Services</a></li>
                <li><a href="#about">About Us</a></li>
                <li><a href="#contact">Contact Us</a></li>
            </ul>
        </div>
    </nav>
  <div class="SameDiv">
    <div class="left-div">
        <img src="./Layout.jpg" height="500px" width="600px" alt="Description of image" class="left-image">
        <div class="content">
            <p> ...........R -> RACK</p>
            <p> ...........MS-R -> MS RACK</p>
            <p> ...........MDR -> MODULAR RACK</p>
            <p> ...........MOD -> MODULAR RACK</p>
            <p> ...........WR -> WIRE RACK</p>
            <p> ...........ALM -> ALMIRAH</p>
        </div>
    </div>

    <div class="right-div">
        <h1>Welcome to Tata Motors</h1>
        <p>Explore our range of vehicles and services.</p>
    </div>
  </div>

    <div class="tools">
        <h2>TOOLS AND MACHINERY LOCATION</h2>
    </div>
    
    <div class="container2">
      <form method="post">
        <input type="numeric" placeholder="Enter the material number" name="search">
        <button class="btn" name="submit">Search</button>
      </form>
      <div class="container3">
        <table class="table">
          <?php
  if(isset($_POST['submit'])){
    $search=$_POST['search'];

    $sql="Select * from `materials` where material_number='$search'";
    $result=mysqli_query($con,$sql);
    if($result){
    if(mysqli_num_rows($result)>0);
      echo '<thead>
        <tr>
        <th>material_number</th>
        <th>location</th>
        <th>material_description</th>
        </tr>
        </thead>';

        $row=mysqli_fetch_assoc($result);
        echo '<tbody>
        <tr>
        <td>'.$row['material_number'].'</td>
        <td>'.$row['location'].'</td>
        <td>'.$row['material_description'].'</td>
        </tr>
        </tbody>';     
    }
  }
  ?>    
        </table>
    </div>
  </body>
</html>