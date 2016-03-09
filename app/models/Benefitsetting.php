<?php

class Benefitsetting extends \Eloquent {
/*
	use \Traits\Encryptable;


	protected $encryptable = [

		'allowance_name',
	];
	*/

public static $rules = [
		'name' => 'required'
	];

	public static $messsages = array(
        'name.required'=>'Please insert benefit name!',
    );

	// Don't forget to fill this array
	protected $fillable = [];


	public function employees(){

		return $this->hasMany('Employee');
	}

}