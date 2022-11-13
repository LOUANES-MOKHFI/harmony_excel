<!DOCTYPE html>
<html lang="{{ str_replace('_', '-', app()->getLocale()) }}">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <title>Laravel</title>

        <!-- Fonts -->
        <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700&display=swap" rel="stylesheet">
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <!-- Styles -->
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.css">
        
        <style>
            html{line-height:1.15;-webkit-text-size-adjust:100%}
            body{margin:0}a{background-color:transparent}[hidden]{display:none}
            html{font-family:system-ui,-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica Neue,Arial,Noto Sans,sans-serif,Apple Color Emoji,Segoe UI Emoji,Segoe UI Symbol,Noto Color Emoji;line-height:1.5}*,:after,:before{box-sizing:border-box;border:0 solid #e2e8f0}a{color:inherit;text-decoration:inherit}svg,video{display:block;vertical-align:middle}video{max-width:100%;height:auto}.bg-white{--bg-opacity:1;background-color:#fff;background-color:rgba(255,255,255,var(--bg-opacity))}.bg-gray-100{--bg-opacity:1;background-color:#f7fafc;background-color:rgba(247,250,252,var(--bg-opacity))}.border-gray-200{--border-opacity:1;border-color:#edf2f7;border-color:rgba(237,242,247,var(--border-opacity))}.border-t{border-top-width:1px}.flex{display:flex}.grid{display:grid}.hidden{display:none}.items-center{align-items:center}.justify-center{justify-content:center}.font-semibold{font-weight:600}.h-5{height:1.25rem}.h-8{height:2rem}.h-16{height:4rem}.text-sm{font-size:.875rem}.text-lg{font-size:1.125rem}.leading-7{line-height:1.75rem}.mx-auto{margin-left:auto;margin-right:auto}.ml-1{margin-left:.25rem}.mt-2{margin-top:.5rem}.mr-2{margin-right:.5rem}.ml-2{margin-left:.5rem}.mt-4{margin-top:1rem}.ml-4{margin-left:1rem}.mt-8{margin-top:2rem}.ml-12{margin-left:3rem}.-mt-px{margin-top:-1px}.max-w-6xl{max-width:72rem}.min-h-screen{min-height:100vh}.overflow-hidden{overflow:hidden}.p-6{padding:1.5rem}
            .py-4{padding-top:1rem;padding-bottom:1rem}.px-6{padding-left:1.5rem;padding-right:1.5rem}
            .pt-8{padding-top:2rem}
            .fixed{position:fixed}
            .relative{position:relative}
            .top-0{top:0}
            .right-0{right:0}
            .shadow{box-shadow:0 1px 3px 0 rgba(0,0,0,.1),0 1px 2px 0 rgba(0,0,0,.06)}
            .text-center{text-align:center}
            
            body {
                font-family: 'Nunito', sans-serif;
            }
            #downloadFile {
                font-family: Arial, Helvetica, sans-serif;
                border-collapse: collapse;
                width: 100%;
                }

                #downloadFile td, #downloadFile th {
                border: 1px solid #ddd;
                padding: 8px;
                }

                #downloadFile tr:nth-child(even){background-color: #f2f2f2;}

                #downloadFile tr:hover {background-color: #ddd;}

                #downloadFile th {
                padding-top: 12px;
                padding-bottom: 12px;
                text-align: left;
                background-color: #04AA6D;
                color: white;
                }
        </style>
    </head>
    <body class="antialiased">
        @include('includes.navBar')
        <div class="relative  items-top justify-center min-h-screen bg-gray-100 dark:bg-gray-900 sm:items-center py-4 sm:pt-0">
            @if (Route::has('login'))
                <div class="hidden fixed top-0 right-0 px-6 py-4 sm:block">
                    @auth
                        <a href="{{ url('/home') }}" class="text-sm text-gray-700 dark:text-gray-500 underline">Home</a>
                    @else
                        <a href="{{ route('login') }}" class="text-sm text-gray-700 dark:text-gray-500 underline">Log in</a>

                        @if (Route::has('register'))
                            <a href="{{ route('register') }}" class="ml-4 text-sm text-gray-700 dark:text-gray-500 underline">Register</a>
                        @endif
                    @endauth
                </div>
            @endif

            <div class="max-w-6xl mx-auto sm:px-6 lg:px-8">
                <div id="loading-image" style="display: none;margin-top:250px" class="text-center">
                    <span class="">
                    <img src="{{asset('loader.gif')}}" style="height: 30px;width: 30px;"> LOADING
                    </span>
                </div> 
                <div class="form">
                    <h3>Exportation</h3>
                    <p>Exporter la table en Excel</p>
                    @if(session()->has('success'))
                                <div class="alert alert-success text-center" id="msg">
                                {{ session()->get('success') }}
                                </div>
                    @elseif(session()->has('error'))
                                <div class="alert alert-danger text-center" id="msg">
                                {{ session()->get('error') }}
                                </div>
                    @endif
                    

                    <form method="POST" id="form" enctype="multipart/form-data">
                        @csrf
                        <div class="row">
                            <div class="col-md-4">
                                <label for="">Insérer le fichier "Ce" à Modifier :</label>
                                <input type="file" name="Ce_file" class="form-control">
                            </div>
                            <div class="col-md-4">
                                <label for="">Insérer le fichier "Fidélis" à modifier :</label>
                                <input type="file" name="fedilis_file" class="form-control">
                            </div>
                            <div class="col-md-4">
                                <label for="">Insérer le fichier "Reporting Agent" à modifier :</label>
                                <input type="file" name="agent_file" class="form-control">
                            </div>
                            <div class="col-md-6">
                                <label for="">Insérer le nouveau fichier d'éxportation:</label>
                                <input type="file" name="fichier" class="form-control">

                            </div>
                            <div class="col-md-6">
                                <label for="">Insérer le nouveau fichier de journal du connexion</label>
                                <input type="file" name="fichierHour" class="form-control">

                            </div>
                            <!--div class="col-md-4">
                                
                                <select class="form-control" name="table" id="table">
                                    <option value="">Choisir une table</option>
                                    @isset($tables)
                                    @foreach($tables as $table)
                                    <option value="{{$table->Tables_in_comunik_excel}}">{{$table->Tables_in_comunik_excel}}</option>
                                    @endforeach
                                    @endisset
                                </select>
                            </div-->
                            <!--div class="col-md-4">
                                <input class="form-control" type="text" name="name" placeholder="Nom de fichier" >
                            </div-->
                            
                            <input type="hidden" name="extension" value="xlsx">
                                <!--select class="form-control" name="extension" >
                                    <option value="xlsx" >.xlsx</option>
                                    <option value="csv" >.csv</option>
                                </select-->
                           
                            <!--div class="col-md-4">
                                <select class="form-control" name="type_document">
                                    <option value="doc1">Document 1</option>
                                    <option value="doc2">Document 2</option>
                                    <option value="doc3">Document 3</option>
                                    <option value="doc4">Document 4</option>
                                </select>
                            </div-->
                        </div>
                        <br>
                        <button type="submit" target="_blank" class="btn btn-info">Exporter</button>
                        <br>
                    </form>

                    <div class="container" id="downloadFile" style="display:none">
                    <br>
                        <table class="table table-bordered table-responsive">
                            <thead>
                                <th>#</th>
                                <th>Fichier</th>
                                <th>Actions</th>
                            </thead>
                            <tbody class="Files">

                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
        <script src="//cdn.jsdelivr.net/npm/sweetalert2@11"></script>
<script>
        $('#form').on('submit',function(e){
        $('.form').hide();
        $('#loading-image').show();
        $('#downloadFile').css('display','none');
        $('.Files').empty();
        e.preventDefault();
        var formData = new FormData(this);
        //let file = $('#file').val();
        /*var file_data = $('#file').prop('files')[0];   
        //alert(file_data);
        var form_data = new FormData();                  
        form_data.append('file', file_data);*/
        
        $.ajax({
                url: '{{ route('excel.editFile') }}',
                type: "POST",
                data:formData,
                cache:false,
                contentType: false,
                processData: false,
                error:function(response){
                    console.log(response);
                },
                success:function(response)
                {           
                    
                    if(response.status == 500){
                        Swal.fire({
                            position: 'center',
                            icon: 'error',
                            title:'error',
                            text: response.msg,
                            showConfirmButton: true,
                            //timer: 5000
                            }
                        );
                    }else{
                        Swal.fire({
                            position: 'center',
                            icon: 'success',
                            title:'success',
                            text: response.msg,
                            showConfirmButton: true,
                           // timer: 5000
                            }
                        );

                        $('#downloadFile').css('display','block');
                        $('.Files').append(`
                                    <tr>
                                        <td>1</td>
                                        <td>Fichier CE_CENTRE</td>
                                        <td><a href="${response.ce}" class="btn btn-success"><i class="fa fa-download"></i>Télècharger</a></td>
                                    </tr>
                                    <tr>
                                        <td>2</td>
                                        <td>Fichier Fidelis_Unadev</td>
                                        <td><a href="${response.fedelis}" class="btn btn-success"><i class="fa fa-download"></i>Télècharger</a></td>

                                    </tr>
                                    <tr>
                                        <td>3</td>
                                        <td>Fichier Reporting Reporting Agents_CALL1</td>
                                        <td><a href="${response.agent}" class="btn btn-success"><i class="fa fa-download"></i>Télècharger</a></td>

                                    </tr>
                                    
                                `);
                                

                    }
                },
                complete: function(){
                    
                    $('#loading-image').hide();
                    $('.form').show();
                    $('#file').empty();
                }
            });
    });
    </script>
    </body>
</html>
