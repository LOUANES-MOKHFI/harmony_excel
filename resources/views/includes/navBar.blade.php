<style>
.topnav {
  overflow: hidden;
  background-color: #333;
}

.topnav a {
  float: left;
  color: #f2f2f2;
  text-align: center;
  padding: 14px 16px;
  text-decoration: none;
  font-size: 17px;
}

.topnav a:hover {
  background-color: #ddd;
  color: black;
}

.topnav a.active {
  background-color: #04AA6D;
  color: white;
}
</style>
</head>
<body>

<div class="topnav">
  <a @if(\Request::route()->getName() == 'home') class="active" @endif href="{{route('home')}}">Call1</a>
  <a @if(\Request::route()->getName() == 'call2') class="active" @endif href="{{route('call2')}}">Call2</a>
  <a @if(\Request::route()->getName() == 'records') class="active" @endif href="{{route('records')}}">Records CRM</a>
  <a @if(\Request::route()->getName() == 'gsm_mail') class="active" @endif href="{{route('gsm_mail')}}">GSM & Mails</a>
  <a @if(\Request::route()->getName() == 'fax_mobile') class="active" @endif href="{{route('fax_mobile')}}">FAX & Mobile</a>
  <a @if(\Request::route()->getName() == 'vicidial_index') class="active" @endif href="{{route('vicidial_index')}}">Vicidial Unadev</a>
  
  <a class="nav-link" href="{{ route('signout') }}">Logout</a>
</div>