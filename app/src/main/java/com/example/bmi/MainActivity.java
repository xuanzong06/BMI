package com.example.bmi;

import android.content.Intent;
import android.provider.MediaStore;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;

public class MainActivity extends AppCompatActivity {

    private Button btn1;
    private Button cambtn;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        FindView();
        btn1.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                go();
            }
        });
        cambtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                cameraup();
            }
        });
    }

    private void FindView() {

        cambtn = findViewById(R.id.cambtn);
        btn1 = findViewById(R.id.gopage2btn);

    }

    public void go(){
        Intent nextpage = new Intent(this, CalculationMBI.class);
        startActivity(nextpage);
    }

    public void cameraup(){
        Intent camerago = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
        startActivity(camerago);
    }
}
