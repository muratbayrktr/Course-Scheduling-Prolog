5 { number(X) : X = 1..20 } 5.


% Ensure exactly 5 numbers are selected.
:- not 5 { number(X) }.

% % Calculate the sum of selected numbers.
sum(S) :- S = #sum { X : number(X) }.

% % Maximize the sum. Penalize solutions with lower sums.
:~ number(X). [-X]

% % Display the selected numbers and the sum.
#show number/1.
#show sum/1.
