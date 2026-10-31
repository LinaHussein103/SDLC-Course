// 1. A variable to store the count
let count = 0;

// 2. Find the elements on the page by their id
const countDisplay = document.getElementById("count");
const message = document.getElementById("message");
const clickBtn = document.getElementById("clickBtn");
const resetBtn = document.getElementById("resetBtn");

// 3. Update what the page shows
function updateDisplay() {
  countDisplay.textContent = count;

  // Bonus: change color above 5
  if (count > 5) {
    countDisplay.style.color = "tomato";
  } else {
    countDisplay.style.color = "black";
  }

  // Bonus: message at 10
  if (count === 10) {
    message.textContent = "You reached 10!";
  } else {
    message.textContent = "";
  }
}

// 4. Click button: add 1
clickBtn.addEventListener("click", function () {
  count = count + 1;
  updateDisplay();
});

// 5. Reset button: back to 0
resetBtn.addEventListener("click", function () {
  count = 0;
  updateDisplay();
});
