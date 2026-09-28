var _dataStart = '';
var _dataEnd = '';

var mutationObserver = new MutationObserver(function (mutations) {
    //var mutationsCount = mutations.length;
    mutations.forEach(function (mutation) {
        if (mutation.type == 'childList'
            && mutation.target.nodeName == 'MSREADOUTSPAN'
            && mutation.target.className == 'msreadout-word-highlight'
            && mutation.addedNodes.length > 0
        )
        {
            //console.log(mutationsCount + " : " + mutation);
            //console.log(mutation);
            //console.log(mutation.target);

            for (let node of mutation.addedNodes) {
                console.log(node.textContent + " : " + mutation.target.className);

                var narrationElement = findAncestor(node, '[data-type="narration"]')

                if (narrationElement) {
                    var dataStart = narrationElement.getAttribute('data-start');
                    var dataEnd = narrationElement.getAttribute('data-end');

                    if (_dataStart != dataStart) {
                        _dataStart = dataStart;
                        _dataEnd = dataEnd;
                        video1.currentTime = dataEnd;
                        console.log(_dataStart + " : " + _dataEnd);
                    }
                }
            }
        }
    });
});

mutationObserver.observe(document.documentElement, {
    attributes: true,
    characterData: true,
    childList: true,
    subtree: true,
    attributeOldValue: true,
    characterDataOldValue: true,
});

function findAncestor(el, sel) {
    while ((el = el.parentElement) && !((el.matches || el.matchesSelector).call(el, sel)));
    return el;
}