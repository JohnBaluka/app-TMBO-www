// This function initializes Reveal.js
export class RevealIO {
	static initializeReveal() {
		// var revealElements = document.querySelectorAll('.reveal');

		// var reveal0 = new Reveal(revealElements[0]).initialize({
		// 	controls: true,
		// 	progress: true,
		// 	hash: true,

		// 	// Learn about plugins: https://revealjs.com/plugins/
		// 	plugins: [RevealZoom, RevealNotes, RevealSearch, RevealMarkdown, RevealHighlight]
		// });

		// var reveal1 = new Reveal(revealElements[1]).initialize({
		// 	controls: true,
		// 	progress: true,
		// 	hash: true,

		// 	// Learn about plugins: https://revealjs.com/plugins/
		// 	plugins: [RevealZoom, RevealNotes, RevealSearch, RevealMarkdown, RevealHighlight]
		// });

		// Also available as an ES module, see:
		// https://revealjs.com/initialization/
		Reveal.initialize({
			controls: true,
			progress: true,
			hash: true,
			width: "1280",
			height: "800",
			showNotes: true,
			slideNumber: "c/t",

			// Learn about plugins: https://revealjs.com/plugins/
			plugins: [RevealZoom, RevealNotes, RevealSearch, RevealMarkdown, RevealHighlight]
		});
	}
}

window.RevealIO = RevealIO;