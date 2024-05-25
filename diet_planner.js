function calculateBMI(weight, height) {
    const heightInMeters = height / 100;
    return (weight / (heightInMeters * heightInMeters)).toFixed(2);
}

function getDietPlan(bmi) {
    // Example diet plan based on BMI
    if (bmi < 18.5) {
        return [
            { day: 'Sunday', breakfast: '2 egg brown bread sandwich + green chutney + 1 cup milk + 3 cashews + 4 almonds + 2 walnuts', midMeal: '1 cup banana shake', lunch: '1 cup achar dal + 1 cup potato curry + 3 chapatti + 1/2 cup rice + 1/2 cup low fat curd + salad', evening: '1 cup strawberry smoothie + 1 cup vegetable poha', dinner: '1.5 cup chicken curry + 3 chapatti + salad' },
            { day: 'Monday', breakfast: '3 onion stuffed paratha + 1 cup curd + 3 cashews + 4 almonds + 2 walnuts', midMeal: '1 cup mango shake', lunch: '1 cup moong dal/ chicken curry + 1 cup potato and cauliflower vegetable + 3 chapatti + 1/2 cup rice + salad', evening: '1 cup pomegranate juice + 2 butter toasted bread', dinner: '1 cup beans potato vegetable + 3 chapatti + salad' },
            { day: 'Tuesday', breakfast: '3 paneer stuffed besan Cheila + green chutney + 1 cup curd + 3 cashews + 4 almonds + 2 walnuts', midMeal: '1 apple smoothie with maple syrup', lunch: '1 cup masoor dal + 1 cup calocasia + 3 chapatti + 1/2 cup rice + 1 cup low curd + salad', evening: '1 cup tomato soup with bread crumbs + 1 cup aloo chaat', dinner: '1 cup carrot peas vegetable +3 chapatti + salad' },
            { day: 'Wednesday', breakfast: '1.5 cup vegetable bread upma + 1 cup milk + 3 cashews + 4 almonds + 2 walnuts', midMeal: '1 cup ripe banana with 2 tsp ghee', lunch: '1 cup rajma curry + 1 cup spinach potato + 3 chapatti + 1/2 cup rice + salad', evening: '1 cup vegetable juice + 1 cup upma', dinner: '1.5 cup Paral vegetable + 3 chapatti + salad' },
            { day: 'Thursday', breakfast: '2 cucumber potato sandwich + 1 tsp green chutney + 1 orange juice + 3 cashews + 2 walnuts + 4 almonds', midMeal: '1 cup buttermilk + 1 cup sweet potato chaat', lunch: '1 cup white chana/ fish curry + 3 chapatti + 1/2 cup rice + salad', evening: '1 cup almond milk + banana', dinner: '1 cup cauliflower potato vegetable + 3 chapatti + salad' },
            { day: 'Friday', breakfast: '2 cup vegetable poha + 1 cup curd + 3 cashews + 4 almonds + 2 walnuts', midMeal: '2 cups watermelon juice', lunch: '1 cup chana dal + 1 cup bhindi vegetable + 3 chapatti + 1/2 cup rice + salad', evening: '1 cup sprouts salad + 2 potato Cheila + green chutney', dinner: '1 cup peas mushroom vegetable + 3 chapatti + salad' },
            { day: 'Saturday', breakfast: '3 vegetable suji Cheila + 1 cup strawberry shake + 4 cashews + 4 almonds + 3 walnuts', midMeal: '1 cup coconut water + 1 cup pomegranate', lunch: '1 cup mix dal + 1 cup soybean curry + 3 chapatti + 1/2 cup curd + salad', evening: '1 cup fruit salad + 4 pc vegetable cutlets + green chutney', dinner: '1 cup karela vegetable + 3 chapati + salad' }
        ];
    } else if (bmi >= 18.5 && bmi < 24.9) {
        return [
            { day: 'Sunday', breakfast: '1 cup vegetable upma + 1/2 cup low-fat milk (no sugar)', midMeal: '1 small apple', lunch: '1 cup chickpea curry + 1 small portion of brown rice + mixed green salad with olive oil and lemon dressing', evening: '1 cup low-fat yogurt with a handful of nuts', dinner: 'Grilled paneer or tofu + steamed broccoli and carrots + 1 whole wheat chapatti' },
            { day: 'Monday', breakfast: 'Smoothie with spinach, banana, low-fat yogurt, and a tablespoon of flax seeds', midMeal: '1 small handful of almonds', lunch: '1 cup moong dal + 1 chapatti + side of mixed vegetable salad', evening: '1 small orange or a handful of berries', dinner: 'Baked fish or tofu + steamed asparagus + 1 small portion of brown rice' },
            { day: 'Tuesday', breakfast: '2 besan cheelas (gram flour pancakes) + 1/2 cup low-fat curd', midMeal: '1 small handful of walnuts', lunch: '1 cup vegetable biryani + 1/2 cup raita + mixed salad', evening: '1 cup tomato soup', dinner: 'Stir-fried tofu with mixed vegetables + 1 chapatti or small portion of brown rice' },
            { day: 'Wednesday', breakfast: '1 bowl vegetable poha + 1/2 cup skim milk (no sugar)', midMeal: '1 small banana', lunch: '1 cup rajma (kidney bean curry) + 1 chapatti + side of mixed salad', evening: '1 cup buttermilk', dinner: 'Grilled chicken or paneer + steamed vegetables + 1 chapatti' },
            { day: 'Thursday', breakfast: '1 cucumber sandwich (whole grain bread) + 1/2 tsp green chutney + 1 apple', midMeal: '1 cup buttermilk', lunch: '1 cup fish curry or tofu curry + 1 chapatti + salad', evening: '1 cup low-fat milk (no sugar)', dinner: '1 cup cabbage and peas vegetable + 1 chapatti + salad' },
            { day: 'Friday', breakfast: '1 bowl vegetable poha (flattened rice) + 1 cup low-fat curd', midMeal: '1 cup watermelon', lunch: '1 cup toor dal (split pigeon peas) + 1 chapatti + salad', evening: '1 cup sprouts salad', dinner: '1 cup pumpkin sabzi (vegetable) + 1 chapatti + salad' },
            { day: 'Saturday', breakfast: '1 cup low-fat milk with oats + 1/2 cup mixed berries', midMeal: '1 cup coconut water', lunch: '1 cup soybean curry + 1 chapatti + 1/2 cup low-fat curd + salad', evening: '1 cup fruit salad', dinner: '1 cup ghia (bottle gourd) vegetable + 1 chapatti + salad' } 
    ]
    } else if (bmi >=25 && bmi<29.9) {
        return [
            { day: 'Sunday', breakfast: '1 bowl oatmeal with 1/2 cup skim milk + 1 banana', midMeal: '1 apple', lunch: '1 cup dal (lentil curry) + 1 chapatti + 1/2 cup low-fat curd + salad', evening: '1 cup vegetable soup', dinner: '1 cup mixed vegetable stir-fry + 1 chapatti + salad' },
            { day: 'Monday', breakfast: '1 stuffed vegetable paratha (without oil) + 1/2 cup low-fat curd', midMeal: '1 cup coconut water', lunch: '1 cup rajma (kidney beans) curry + 1 chapatti + salad', evening: '1 cup mixed fruit salad', dinner: '1 cup palak paneer (spinach with cottage cheese) + 1 chapatti + salad' },
            { day: 'Tuesday', breakfast: '2 besan cheelas (gram flour pancakes) + 1/2 cup low-fat curd', midMeal: '1 orange', lunch: '1 cup chana masala (chickpeas) + 1 chapatti + salad', evening: '1 cup tomato soup', dinner: '1 cup baingan bharta (mashed eggplant) + 1 chapatti + salad' },
            { day: 'Wednesday', breakfast: '1 bowl vegetable upma + 1/2 cup skim milk (no sugar)', midMeal: '1 pear', lunch: '1 cup moong dal (green gram curry) + 1 chapatti + salad', evening: '1 cup vegetable soup', dinner: '1 cup methi thepla (fenugreek flatbread) + salad' },
            { day: 'Thursday', breakfast: '1 cucumber sandwich (whole grain bread) + 1/2 tsp green chutney + 1 apple', midMeal: ' 1 cup buttermilk', lunch: '1 cup fish curry or tofu curry + 1 chapatti + salad', evening: '1 cup low-fat milk (no sugar)', dinner: '1 cup cabbage and peas vegetable + 1 chapatti + salad' },
            { day: 'Friday', breakfast: '1 bowl vegetable poha (flattened rice) + 1 cup low-fat curd', midMeal: '1 cup watermelon', lunch: '1 cup toor dal (split pigeon peas) + 1 chapatti + salad', evening: ' 1 cup sprouts salad', dinner: '1 cup pumpkin sabzi (vegetable) + 1 chapatti + salad' },
            { day: 'Saturday', breakfast: '1 cup low-fat milk with oats + 1/2 cup mixed berries', midMeal: '1 cup coconut water', lunch: '1 cup soybean curry + 1 chapatti + 1/2 cup low-fat curd + salad', evening: '1 cup fruit salad', dinner: '1 cup ghia (bottle gourd) vegetable + 1 chapatti + salad' }
        ];
    } else if (bmi >=30) {
        return [
            { day: 'Sunday', breakfast: '3 egg whites + 1 toasted brown bread + 1/2 cup low-fat milk (no sugar)', midMeal: '1 cup papaya', lunch: '1 cup arhar dal + 1 chapatti + 1/2 cup low-fat curd + salad', evening: '1 cup vegetable soup', dinner: '1 cup pumpkin + 1 chapatti + salad' },
            { day: 'Monday', breakfast: '1 onion stuffed chapatti + 1/2 cup low-fat curd', midMeal: '1 cup coconut water', lunch: '1 cup moong dal/ chicken curry + 1 chapatti + salad', evening: '1 cup pomegranate', dinner: '1 cup beans + 1 chapatti + salad' },
            { day: 'Tuesday', breakfast: '2 besan cheela + 1/2 cup low-fat curd', midMeal: '1 apple', lunch: '1 cup masoor dal + 1 chapatti + 1/2 cup low-fat curd + salad', evening: '1 cup tomato soup', dinner: '1 cup carrot peas vegetable + 1 chapatti + salad' },
            { day: 'Wednesday', breakfast: '1 cup vegetable brown bread upma + 1/2 cup low-fat milk (no sugar)', midMeal: '1 cup musk melon', lunch: '1 cup rajma curry + 1 chapatti + salad', evening: '1 cup vegetable soup', dinner: '1 cup parwal vegetable + 1 chapatti + salad' },
            { day: 'Thursday', breakfast: '1 cucumber hung curd sandwich + 1/2 tsp green chutney + 1 orange', midMeal: '1 cup buttermilk', lunch: '1 cup white chana/ fish curry + 1 chapatti + salad', evening: '1 cup low-fat milk (no sugar)', dinner: '1 cup cauliflower vegetable + 1 chapatti + salad' },
            { day: 'Friday', breakfast: '1 cup vegetable poha + 1 cup low-fat curd', midMeal: '1 cup watermelon', lunch: '1 cup chana dal + 1 chapatti + salad', evening: '1 cup sprouts salad', dinner: '1 cup tinda vegetable + 1 chapatti + salad' },
            { day: 'Saturday', breakfast: '1 cup low-fat milk with oats + 3-4 strawberries', midMeal: '1 cup coconut water', lunch: '1 cup soybean curry + 1 chapatti + 1/2 cup low-fat curd + salad', evening: '1 cup fruit salad', dinner: '1 cup ghia vegetable + 1 chapatti + salad' }
        ];
    }
}

function getPrecautions(bmi) {
    if (bmi < 18.5) {
        return {
            dos: [
                "Eat nutrient-dense foods high in vitamins, minerals, and healthy fats.",
                "Include protein-rich foods like lean meats, fish, eggs, beans, and legumes.",
                "Eat small, frequent meals to increase calorie intake.",
                "Stay hydrated by drinking plenty of water.",
                "Choose healthy fats such as avocado, nuts, seeds, and olive oil.",
                "Monitor physical activity with light to moderate exercise.",
                "Consult a healthcare provider for personalized advice.",
                "Add calorie-dense snacks like nuts, dried fruits, and yogurt.",
                "Consider nutritional supplements under guidance.",
                "Track progress by monitoring weight gain."
            ],
            donts: [
                "Avoid relying solely on junk food for empty calories.",
                "Don't skip meals; regular eating is essential.",
                "Avoid over-exercising, which can burn more calories than you consume.",
                "Don't restrict major food groups; balanced nutrition is key.",
                "Don't ignore symptoms; consult a healthcare professional.",
                "Eat when hungry to prevent unintentional weight loss.",
                "Avoid comparing yourself to others for healthy gain.",
                "Limit excessive caffeine consumption to avoid appetite suppression.",
                "Seek support for body image concerns.",
                "Remember that weight gain requires patience and consistency."
            ],
            precautions: [
                "Monitor bone health and consider calcium and vitamin D supplements if necessary.",
                "Focus on strength training instead of excessive cardio.",
                "Regularly check for nutrient deficiencies like iron, B vitamins, and zinc.",
                "Consult with a dietitian to ensure healthy weight gain.",
                "Monitor mental health and seek support if needed.",
                "Avoid extreme dietary restrictions for balanced nutrition.",
                "Consider medical evaluation for underlying conditions.",
                "Take breaks during physical activity to prevent overexertion.",
                "Use supplements cautiously under medical supervision.",
                "Maintain a regular sleep schedule for overall health."
            ]
        };
    } else if (bmi >= 18.5 && bmi < 24.9) {
        return {
            dos: [
                "Eat a balanced diet with fruits, vegetables, lean proteins, and whole grains.",
                "Practice portion control to maintain weight.",
                "Engage in regular exercise for overall health.",
                "Drink plenty of water to stay hydrated.",
                "Eat mindfully, paying attention to hunger and fullness cues.",
                "Get regular health check-ups for monitoring.",
                "Include healthy fats like avocado, nuts, seeds, and fish.",
                "Limit processed foods and sugary drinks.",
                "Manage stress with relaxation techniques.",
                "Adjust diet and activity based on energy levels."
            ],
            donts: [
                "Avoid overeating unhealthy, high-calorie foods.",
                "Don't live a sedentary lifestyle; aim for regular physical activity.",
                "Ensure each meal is balanced; don't skip meals.",
                "Maintain hydration with water, not sugary drinks.",
                "Avoid comparing yourself to others.",
                "Seek support for stress or emotional eating.",
                "Limit alcohol intake for overall health.",
                "Aim for adequate sleep to support health.",
                "Stay active daily; don't ignore physical activity.",
                "Avoid excessive screen time for reduced sedentary behavior."
            ],
            precautions: [
                "Manage stress levels to support overall health.",
                "Be cautious with portion sizes to avoid overeating.",
                "Limit alcohol consumption for overall health.",
                "Enjoy treats in moderation without overindulging.",
                "Maintain regular physical activity habits.",
                "Stay hydrated throughout the day.",
                "Limit screen time to reduce sedentary behavior.",
                "Monitor cholesterol levels for heart health.",
                "Report significant changes in eating habits to a healthcare provider.",
                "Verify health claims before adopting new dietary practices."
            ]
        };
    } else if (bmi >= 25 && bmi < 29.9) {
        return {
            dos: [
                "Eat a balanced diet with fruits, vegetables, lean proteins, and whole grains.",
                "Control portion sizes to manage calorie intake.",
                "Engage in at least 150 minutes of moderate-intensity exercise per week.",
                "Include fiber-rich foods like beans, lentils, and vegetables.",
                "Stay hydrated to support metabolism.",
                "Monitor weight changes and adjust diet and activity levels.",
                "Consider working with a dietitian or healthcare provider.",
                "Reduce consumption of sugary drinks.",
                "Eat mindfully to avoid overeating.",
                "Set realistic goals for sustainable changes."
            ],
            donts: [
                "Minimize intake of processed foods and sugary snacks.",
                "Maintain regular eating patterns; don't skip meals.",
                "Stay consistent with physical activity for weight management.",
                "Avoid binge eating; consume moderate portions.",
                "Aim for adequate sleep for health and metabolism.",
                "Focus on long-term lifestyle changes, not quick fixes.",
                "Address emotional triggers for healthy eating habits.",
                "Incorporate resistance exercises; don't neglect strength training.",
                "Drink water instead of sugary beverages.",
                "Focus on personal progress, not comparisons."
            ],
            precautions: [
                "Monitor blood pressure to manage risks associated with overweight status.",
                "Maintain regular exercise for weight management and health.",
                "Avoid crash diets for sustainable weight loss.",
                "Address emotional triggers for overeating.",
                "Consider low-impact exercises for joint health.",
                "Track weight changes and adjust diet and exercise.",
                "Minimize processed snacks and high-sugar foods.",
                "Seek guidance from healthcare professionals.",
                "Discuss symptoms of sleep apnea with a healthcare provider.",
                "Focus on balanced nutrition for sustainable weight loss."
            ]
        };
    } else if (bmi >= 30) {
        return {
            dos: [
                "Seek professional guidance for personalized weight management.",
                "Adopt a balanced diet with fruits, vegetables, lean proteins, and whole grains.",
                "Control portion sizes to manage calorie intake.",
                "Engage in at least 150 minutes of moderate-intensity exercise per week, as recommended.",
                "Monitor health indicators like blood pressure and cholesterol levels.",
                "Include fiber-rich foods for digestive health and satiety.",
                "Stay hydrated throughout the day.",
                "Set realistic goals for long-term changes in diet and exercise habits.",
                "Seek support from social networks or groups.",
                "Celebrate achievements in health improvement."
            ],
            donts: [
                "Follow healthcare provider recommendations for weight management.",
                "Avoid fad diets promising rapid, unsustainable weight loss.",
                "Maintain regular eating patterns; don't skip meals.",
                "Stay consistent with physical activity for weight loss.",
                "Avoid emotional eating triggers; address mental health.",
                "Limit alcohol intake contributing to excess calories.",
                "Focus on personal progress, not comparisons.",
                "Drink water, not sugary beverages.",
                "Stay committed to health goals despite setbacks.",
                "Avoid extreme calorie restriction for sustainable weight loss."
            ],
            precautions: [
                "Monitor cardiovascular health to manage risks associated with obesity.",
                "Address metabolic syndrome risks like high blood sugar with medical guidance.",
                "Schedule regular medical check-ups for overall health.",
                "Develop a comprehensive weight management plan with healthcare providers.",
                "Discuss symptoms of sleep disorders with a healthcare provider.",
                "Maintain regular exercise for weight management.",
                "Monitor liver function to prevent fatty liver disease.",
                "Seek support for mental well-being impacted by obesity.",
                "Consider risks and benefits of weight loss surgery with medical guidance.",
                "Focus on balanced nutrition for sustainable weight loss."
            ]
        };
    }
}


function generateDietPlan() {
    const name = document.getElementById('name').value;
    const age = document.getElementById('age').value;
    const gender = document.querySelector('input[name="gender"]:checked');
    const height = parseFloat(document.getElementById('height').value);
    const weight = parseFloat(document.getElementById('weight').value);
    const bmi = calculateBMI(weight, height);

    if (isNaN(height) || isNaN(weight)) {
        alert("Please enter valid height and weight values.");
        return;
    }

    const dietPlan = getDietPlan(bmi);
    const precautions = getPrecautions(bmi) || { dos: [], donts: [], precautions: [] };

    let dietHTML = `<h2>Diet Plan</h2>`;
    dietHTML += `<table border="1"><tr><th>Day</th><th>Breakfast</th><th>Mid-Meal</th><th>Lunch</th><th>Evening</th><th>Dinner</th></tr>`;
    dietPlan.forEach(day => {
        dietHTML += `<tr>
            <td>${day.day}</td>
            <td>${day.breakfast}</td>
            <td>${day.midMeal}</td>
            <td>${day.lunch}</td>
            <td>${day.evening}</td>
            <td>${day.dinner}</td>
        </tr>`;
    });
    dietHTML += `</table>`;

    dietHTML += `<div class="page-break"></div>
                 <h2>Dos</h2>
                 <ul>`;
    precautions.dos.forEach(doItem => {
        dietHTML += `<li>${doItem}</li>`;
    });
    dietHTML += `</ul>`;

    dietHTML += `<div class="page-break"></div>
                 <h2>Don'ts</h2>
                 <ul>`;
    precautions.donts.forEach(dontItem => {
        dietHTML += `<li>${dontItem}</li>`;
    });
    dietHTML += `</ul>`;

    dietHTML += `<div class="page-break"></div>
                 <h2>Precautions</h2>
                 <ul>`;
    precautions.precautions.forEach(precautionItem => {
        dietHTML += `<li>${precautionItem}</li>`;
    });
    dietHTML += `</ul>`;

    document.getElementById('dietPlan').innerHTML = dietHTML;
    document.getElementById('downloadPdf').style.display = 'block';
}

function generatePDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    const name = document.getElementById('name').value;
    const age = document.getElementById('age').value;
    const gender = document.querySelector('input[name="gender"]:checked').value;
    const height = parseFloat(document.getElementById('height').value);
    const weight = parseFloat(document.getElementById('weight').value);
    const bmi = calculateBMI(weight, height);

    const currentDate = new Date();
    const timestamp = currentDate.toLocaleString();

    doc.setFontSize(12);
    doc.setFont('Times-Roman', 'normal');
    const pageWidth = doc.internal.pageSize.getWidth();
    const timestampXPosition = pageWidth - doc.getTextWidth(`Generated on: ${timestamp}`) - 20;
    const dateXPosition = pageWidth - doc.getTextWidth(`Date: ${currentDate.toDateString()}`) - 20;
    doc.text(`Generated on: ${timestamp}`, timestampXPosition, 40);
    doc.text(`Date: ${currentDate.toDateString()}`, dateXPosition, 30);

    doc.setFontSize(24);
    doc.setFont('Times-Roman','bold');
    doc.text('BMI and Diet Suggestions Report', 50, 20);
    doc.setFont('Times-Roman','normal');
    doc.setFontSize(12);
    doc.text(`Name: ${name}`, 20, 40);
    doc.text(`Age: ${age}`, 20, 50);
    doc.text(`Gender: ${gender}`, 20, 60);
    doc.text(`Height: ${height} cm`, 20, 70);
    doc.text(`Weight: ${weight} kg`, 20, 80);
    doc.text(`BMI: ${bmi}`, 20, 90);

    let bmiCategory = '';
    if (bmi < 18.5) {
        bmiCategory = 'Underweight';
    } else if (bmi >= 18.5 && bmi < 24.9) {
        bmiCategory = 'Normal Weight';
    } else if (bmi >= 25 && bmi < 29.9) {
        bmiCategory = 'Overweight';
    } else {
        bmiCategory = 'Obese';
    }
    doc.text(`BMI Category: ${bmiCategory}`, 20, 100);

    const dietPlan = getDietPlan(bmi);

    doc.setFontSize(14);
    doc.setFont('Times-Roman','bold');
    doc.text('Diet Suggestions:', 20, 110);
    doc.setFont('Times-Roman','normal');
    doc.setFontSize(12);
    doc.text('NOTE: The below are the diet suggestions. Please consult a doctor if you suffer from any diseases.', 20, 115);

    const dietTable = dietPlan.map(day => [day.day, day.breakfast, day.midMeal, day.lunch, day.evening, day.dinner]);

    doc.autoTable({
        head: [['Day', 'Breakfast\n(8:00-8:30AM)', 'Mid-Meal\n(11:00-11:30AM)', 'Lunch\n(2:00-2:30PM)', 'Evening\n(4:00-4:30PM)', 'Dinner\n(8:00-8:30PM)']],
        body: dietTable,
        startY: 120,
        theme: 'grid',
        styles: { fontSize: 10 }
    });
    doc.setDrawColor(0);
    doc.setLineWidth(1);
    doc.setDrawColor(169, 169, 169);
    doc.rect(10, 10,190,277);
    const precautions = getPrecautions(bmi) || { dos: [], donts: [], precautions: [] };

    // Add Dos section
    doc.addPage();
    doc.setDrawColor(0);
    doc.setLineWidth(1);
    doc.setDrawColor(169, 169, 169);
    doc.rect(10, 10,190,277);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.text("Dos :", 20, 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(12);
    let dosPageHeight = 30;
    precautions.dos.forEach((doItem, index) => {
        const textLines = doc.splitTextToSize(`${index + 1}. ${doItem}`, doc.internal.pageSize.width - 40);
        if (dosPageHeight + (textLines.length * 10) > doc.internal.pageSize.height - 20) {
            doc.addPage();
            dosPageHeight = 20;
        }
        textLines.forEach(line => {
            doc.text(line, 20, dosPageHeight);
            dosPageHeight += 10;
        });
    });

    // Add Don'ts section
    doc.addPage();
    doc.setDrawColor(0);
    doc.setLineWidth(1);
    doc.setDrawColor(169, 169, 169);
    doc.rect(10, 10,190,277);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.text("Don'ts :", 20, 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(12);
    precautions.donts.forEach((dontItem, index) => {
        doc.text(`${index + 1}. ${dontItem}`, 20, 30 + (index * 10));
    });

    // Add Precautions section
    doc.addPage();
    doc.setDrawColor(0);
    doc.setLineWidth(1);
    doc.setDrawColor(169, 169, 169);
    doc.rect(10, 10,190,277);
    doc.setFontSize(14);
    doc.setFont("helvetica", "bold");
    doc.text("Precautions :", 20, 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(12);
    precautions.precautions.forEach((precautionItem, index) => {
        doc.text(`${index + 1}. ${precautionItem}`, 20, 30 + (index * 10));
    });

    doc.save('diet_plan.pdf');
}
