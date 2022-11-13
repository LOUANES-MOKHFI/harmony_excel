@extends('master')
@section('title')
    GSMs & Mails
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
            <h3>GSM & MAILS</h3>
            <p>Exporter le fichiers des GSM et Mails</p>
            @if(session()->has('success'))
                        <div class="alert alert-success text-center" id="msg">
                        {{ session()->get('success') }}
                        </div>
            @elseif(session()->has('error'))
                        <div class="alert alert-danger text-center" id="msg">
                        {{ session()->get('error') }}
                        </div>
            @endif
            

            <form  method="POST" id="" action="{{ route('downloadgsm_mail') }}" enctype="multipart/form-data">
                @csrf
                <div class="row">
                    <div class="col-md-4">
                        <label for="">type d'exportation :</label>
                        <select name="type_export" id="" required class="form-control">
                            <option value="" selected> Choisir le type d'exportation</option>
                            <option value="GSM">GSMs</option>
                            <option value="MAIL">Mails</option>
                        </select>
                    </div>
                    <div class="col-md-8">
                        <label for="">Insérer le fichier d'exportation :</label>
                        <input type="file" name="export_file" class="form-control">
                    </div>
                    
                    <input type="hidden" name="extension" value="xlsx">
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
                url: '{{ route('downloadgsm_mail') }}',
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
                                        <td>Fichier Des GSMs modifiés</td>
                                        <td><a href="${response.gsm}" class="btn btn-success"><i class="fa fa-download"></i>Télècharger</a></td>
                                    </tr>
                                    <tr>
                                        <td>2</td>
                                        <td>Fichier Fichier Des Mails modifiés</td>
                                        <td><a href="${response.mails}" class="btn btn-success"><i class="fa fa-download"></i>Télècharger</a></td>

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
        