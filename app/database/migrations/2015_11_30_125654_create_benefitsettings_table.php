<?php

use Illuminate\Database\Schema\Blueprint;
use Illuminate\Database\Migrations\Migration;

class CreateBenefitSettingsTable extends Migration {

	/**
	 * Run the migrations.
	 *
	 * @return void
	 */
	public function up()
	{
		Schema::create('benefitsettings', function(Blueprint $table)
		{
			$table->increments('id');
			$table->string('benefit_name');
			$table->integer('organization_id')->unsigned()->default('0')->index('benefits_organization_id_foreign');
			$table->timestamps();
		});

		
	}


	/**
	 * Reverse the migrations.
	 *
	 * @return void
	 */
	public function down()
	{
		Schema::drop('benefitsettings');
	}

}
