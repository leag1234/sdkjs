/*
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function(window)
{
	window['AscCommon'] = window['AscCommon'] || {};
	var AscCommon = window['AscCommon'];

	function parseNumber(value)
	{
		var parsed = parseFloat(value);
		return isFinite(parsed) ? parsed : 1;
	}

	function getTransformScale(transform)
	{
		var match = /^\s*(matrix|scale)\s*\(([^)]*)\)\s*$/i.exec(transform || '');
		if (!match)
			return 1;

		var values = match[2].split(/[\s,]+/).filter(function(value) { return value.length; });
		if (match[1].toLowerCase() === 'scale')
		{
			var scaleX = parseNumber(values[0]);
			var scaleY = values.length > 1 ? parseNumber(values[1]) : scaleX;
			return Math.sqrt(Math.abs(scaleX * scaleY));
		}

		if (values.length < 6)
			return 1;
		var a = parseNumber(values[0]), b = parseNumber(values[1]);
		var c = parseNumber(values[2]), d = parseNumber(values[3]);
		return Math.sqrt(Math.abs(a * d - b * c));
	}

	AscCommon['EoSvgEffectiveScale'] = function(transformStrings, viewBoxFactor)
	{
		var scale = isFinite(viewBoxFactor) && viewBoxFactor > 0 ? viewBoxFactor : 1;
		for (var index = 0; index < transformStrings.length; index++)
			scale *= getTransformScale(transformStrings[index]);
		return scale;
	};

	AscCommon['EoFixSvgHairlines'] = function(svgText)
	{
		try
		{
			if (typeof svgText !== 'string' || !svgText || typeof DOMParser !== 'function' || typeof XMLSerializer !== 'function')
				return svgText;

			var doc = new DOMParser().parseFromString(svgText, 'image/svg+xml');
			if (!doc || !doc.documentElement || doc.getElementsByTagName('parsererror').length)
				return svgText;

			var root = doc.documentElement;
			var factor = 1;
			var viewBox = (root.getAttribute('viewBox') || '').trim().split(/[\s,]+/);
			var width = parseFloat(root.getAttribute('width'));
			if (viewBox.length === 4 && isFinite(width) && parseFloat(viewBox[2]) > 0)
				factor = width / parseFloat(viewBox[2]);

			var elements = doc.getElementsByTagName('*');
			var changed = false;
			for (var index = 0; index < elements.length; index++)
			{
				var element = elements[index];
				var stroke = element.getAttribute('stroke');
				if (element.style && element.style.stroke)
					stroke = element.style.stroke;
				if (!stroke)
					continue;
				if (!stroke || stroke === 'none' || element.getAttribute('vector-effect'))
					continue;

				var widthValue = element.getAttribute('stroke-width');
				if (element.style && element.style.strokeWidth)
					widthValue = element.style.strokeWidth;
				var strokeWidth = widthValue == null || widthValue === '' ? 1 : parseFloat(widthValue);
				if (!isFinite(strokeWidth) || strokeWidth <= 0)
					continue;

				var transforms = [];
				for (var node = element; node && node.getAttribute; node = node.parentNode)
				{
					var transform = node.getAttribute('transform');
					if (transform)
						transforms.unshift(transform);
				}
				if (strokeWidth * AscCommon.EoSvgEffectiveScale(transforms, factor) >= 1)
					continue;

				element.setAttribute('vector-effect', 'non-scaling-stroke');
				element.setAttribute('stroke-width', '1');
				if (element.style)
					element.style.removeProperty('stroke-width');
				changed = true;
			}
			return changed ? new XMLSerializer().serializeToString(doc) : svgText;
		}
		catch (error)
		{
			return svgText;
		}
	};

	AscCommon['EoFixSvgHairlines'] = AscCommon.EoFixSvgHairlines;
	AscCommon['EoSvgEffectiveScale'] = AscCommon.EoSvgEffectiveScale;
})(window);
