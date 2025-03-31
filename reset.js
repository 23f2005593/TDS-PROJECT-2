inputs = document.querySelectorAll('input')
textboxes = document.querySelectorAll("textarea")
buttons  = document.querySelectorAll("button")
inputs.forEach(input => input.removeAttribute('disabled'));
buttons.forEach(input => input.removeAttribute('disabled'));
textboxes.forEach(input => input.removeAttribute('disabled'));