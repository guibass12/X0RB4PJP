$(function() {
	// Disable text selection & dragging
	RaUtils.preventElementSelection();
	RaUtils.preventElementDrag();

	// Cache images
	RaUtils.cacheImages('img/x-dark.png');

	$('#ph-img-logo').attrchange({
		trackValues: false,
		callback: function(e) {
			RaUtils.replaceImgWithBackImage('#logo', '#ph-img-logo');
		}
	});
	$('#ph-img-image').attrchange({
		trackValues: false,
		callback: function(e) {
			RaUtils.replaceImgWithBackImage('#image', '#ph-img-image');
		}
	});
	$('#ph-btn-color').attrchange({
		trackValues: false,
		callback: function(e) {
			RaUtils.replaceButtonColor('#accept', '#ph-btn-color');
		}
	});

	$(document).on(RaSDK.events.ON_TEMPLATE_READY, function() {
		RaUtils.replaceImgWithBackImage('#logo', '#ph-img-logo');
		RaUtils.replaceImgWithBackImage('#image', '#ph-img-image');
		RaUtils.replaceButtonColor('#accept', '#ph-btn-color');
		RaSDK.open();
	});

	RaSDK.updateTemplate();
});


function DoClose() {
	RaSDK.run('close');
}

function DoInstall() {
	RaSDK.run('install {"button": "accept"}');
	window.setInterval(function() {
		DoClose();
	}, 100);
}