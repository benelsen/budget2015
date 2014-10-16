
generate: budget.xls
	node generate.js

budget.xls:
	curl -o budget.xls http://budget.public.lu/wp-content/uploads/2014/10/Recettes-et-D%C3%A9penses-au-10-octobre-2014.xls

