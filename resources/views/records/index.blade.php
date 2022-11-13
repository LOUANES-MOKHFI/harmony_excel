@extends('master')
@section('title')
    Download Records
@endsection

@section('css')
@endsection
@section('content')

<div class="relative  items-top justify-center min-h-screen bg-gray-100 dark:bg-gray-900 sm:items-center py-4 sm:pt-0">

    <div class="max-w-6xl mx-auto sm:px-6 lg:px-8">
        <div id="loading-image" style="display: none;margin-top:250px" class="text-center">
            <span class="">
            <img src="{{asset('loader.gif')}}" style="height: 30px;width: 30px;"> LOADING
            </span>
        </div> 
        <div class="form">
            <h3>Download Recording From CRM</h3>
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
            

            <form method="POST" id="" action="{{ route('downloadRecords') }}" enctype="multipart/form-data">
                @csrf
                <div class="row">
                    <div class='col-md-4'>
                        <label for="">Choisir le server</label>
                        <select name="serverType" id="" class="form-control" required>
                            <option value="" selected> veuillez choisir le server</option>
                            <option value="CALL1"> CALL 1</option>
                            <option value="CALL2"> CALL 2</option>
                        </select>
                    </div>
                    <div class="col-md-8">
                        <label for="">Insérer le fichier Recordin:</label>
                        <input type="file" name="recordsFile" class="form-control" required>
                    </div>                       
                </div>
                <br>
                <button type="submit" target="_blank" class="btn btn-info">Envoyer</button>
                <br>
            </form>
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
                url: '{{ route('downloadRecords') }}',
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
        