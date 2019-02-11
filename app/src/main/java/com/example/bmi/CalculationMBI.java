package com.example.bmi;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

public class CalculationMBI extends AppCompatActivity {

    private EditText height;
    private EditText weight;
    private Button bmibtn;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_calculation_mbi);

        Findview();
        bmibtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                BMI();
            }
        });
    }

    private void Findview() {
        height = findViewById(R.id.height);
        weight = findViewById(R.id.weight);
        bmibtn = findViewById(R.id.bmibtn);
    }

    public void BMI(){
        String h = height.getText().toString();
        String w = weight.getText().toString();
        float height1 = Float.parseFloat(h);
        float weight1 = Float.parseFloat(w);
        float bmi = weight1 / (height1/100 * height1/100);

        Toast.makeText(this, String.valueOf(bmi), Toast.LENGTH_LONG).show();
    }
}
