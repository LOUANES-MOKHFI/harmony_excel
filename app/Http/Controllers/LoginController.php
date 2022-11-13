<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Hash;
use Session;
use App\Models\User;
use Illuminate\Support\Facades\Auth;
class LoginController extends Controller
{
    public function index()
    {
        return view('auth.login');
    }  
      
    public function customLogin(Request $request)
    {
        $request->validate([
            'email' => 'required',
            'password' => 'required',
        ]);
   
        $credentials = $request->only('email', 'password');
        if (Auth::attempt($credentials)) {
            return redirect()->route('home');
                        //->withSuccess('Signed in');
        }
  
        return  redirect()->route("login")->withError('Email ou mot de passe incorrecte');
    }

    public function signOut() {
        Session::flush();
        Auth::logout();
  
        return redirect()->route("login");
    }
}
