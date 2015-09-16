var timer = -1;

function hidePopup() {
	$(".popup-menu").hide();
	timer = -1;
}

function showPopup(items, x, y) {
	var $popup = $(".popup-menu");
	$popup.empty();
	for (var i = 0; i < items.length; ++i) {
		var item = items[i],
			$listItem = $("<li></li>"),
			$link = $("<a></a>")
				.attr("href", "#" + item.id)
				.text(item.text);
			$listItem.append($link);
			$popup.append($listItem);
	}
	if (timer >= 0) {
		window.clearTimeout(timer);
	}			
	$popup
		.css("position", "absolute")
		.css("left", x)
		.css("top", y)
		.show();
}

$(document).ready(function () {
	// Create a dummy container for the popup menu
	$("<ul></ul>")
		.attr("class", "popup-menu")
		.css("display", "none")
		.on("mouseleave", function() {
			if (timer == -1) {
				timer = window.setTimeout(hidePopup, 1000);
			}
		})
		.on("mouseenter", function() {
			if (timer >= 0) {
				window.clearTimeout(timer);
				timer = -1;
			}
		})
		.insertAfter(".main");
	
	// Add backlinks for foreign keys
	$("H1").each(function () {
		var $h1 = $(this),
			id = $h1.attr("id"),
			links = [];
		if (id) {
			var $links = $("a[href='#" + id + "']");
			$links.each(function () {
				var $a = $(this),
					$container = $a.closest(".table-container"),
					$header = $container.find("h1").filter(":first"),
					refid = $header.attr("id");
				if ($header.length > 0 && !links.some(function (c, i, a) { return c && c.id && c.id === refid; })) {
					links.push( { id: refid, text: $header.text() } );
				}
			});
		}
		if (links.length > 0) {
			$h1.data("refs", links)
				.css("cursor", "pointer")
				.on("click", function (e) {
					var items = $(this).data("refs");
					showPopup(items, e.target.offsetLeft + e.offsetX, e.target.offsetTop + e.offsetY);
				});
		}
	});
});
