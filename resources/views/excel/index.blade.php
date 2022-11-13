@extends('master')
@section('title')
    Call 1
@endsection

@section('css')
@endsection
@section('content')

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
                    <div class='col-md-8'>
                        <label for="">Choisir le jour :</label>
                        <input type="date" name="date" class="form-control"><br>
                    </div>
                    <div class="col-md-4">
                        <label for="">Client :</label>
                        <select class="form-control" name="client_name" required>
                            <option value="" selected> -- Choisir un client -- </option>
                            <option value="unadev"><span class="text-success">UNADEV</span> </option>
                            <option value="unapei"><span class="text-danger">UNAPEI</span> </option>
                        </select>
                    </div>
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
@endsection

@section('js')
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
@endsection
        