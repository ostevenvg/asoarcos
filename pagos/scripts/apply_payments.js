const delay = ms => new Promise(res => setTimeout(res, ms));

function find_btn(tag, text, bclass) {
	var aTags = document.getElementsByTagName(tag);
	var searchText = text;
	var found;

	for (var i = 0; i < aTags.length; i++) {
		if (aTags[i].innerText == searchText &&  (bclass == "ignore" || aTags[i].className ==bclass)) {
	    		found = aTags[i];
	    		break
	  	  }
	}
	return found
}

const apply_all_payments = async () => {
	let i = 0
	while ( i<5 ) {
		i = i + 1
		console.log("Iteración " + i);

		pay_btn = find_btn('strong', 'Receive payment', 'ignore');
		if (pay_btn === undefined) {
			console.log("No mas botones de pago encontrados. Fin del proceso")
			break
		}
		pay_btn.click()
		console.log("Esperando 10 segundos antes de ejecutar Save and close");
    		await delay(10000);
			
		save_btn = find_btn('button', 'Save and close', 'combo-button-main')
		save_btn.click()
		console.log("Esperando 10 segundos antes de buscar otro pago");
	    	await delay(10000);
			
	}
};

apply_all_payments()