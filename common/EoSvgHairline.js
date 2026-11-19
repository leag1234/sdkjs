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

	// Width of a hairline in device pixels. Word writes the lines of a WMF/EMF drawing as hairlines of 0.14 pt in a
	// PDF; on screen they are thinner than a 1 px line. 0.6 px matches the reporter's comparison with the table
	// borders (#61); on high-density screens the width follows the pixel ratio.
	var HAIRLINE_PX = 0.6;
	// 0.14 pt (Word's hairline in a PDF) expressed in pixels of a raster of the given resolution
	var HAIRLINE_PT = 0.14;

	AscCommon['EoSvgHairlinePx'] = function()
	{
		var ratio = AscCommon.AscBrowser && AscCommon.AscBrowser.retinaPixelRatio;
		return HAIRLINE_PX * (isFinite(ratio) && ratio >= 1 ? ratio : 1);
	};

	AscCommon['EoSvgHairlinePxForDpi'] = function(dpi)
	{
		return HAIRLINE_PT * dpi / 72;
	};

	// options.minPx: width of the hairline in destination pixels (default: EoSvgHairlinePx)
	// options.destWidthPx: width in pixels at which the picture will be rasterized (default: the SVG size, scale 1)
	// options.userUnits: write the width in the units of each element instead of vector-effect non-scaling-stroke
	AscCommon['EoFixSvgHairlines'] = function(svgText, options)
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

			var minPx = options && options.minPx > 0 ? options.minPx : AscCommon.EoSvgHairlinePx();
			var destScale = 1;
			if (options && options.destWidthPx > 0)
			{
				var intrinsicWidth = isFinite(width) && width > 0 ? width : parseFloat(viewBox[2]);
				if (isFinite(intrinsicWidth) && intrinsicWidth > 0)
					destScale = options.destWidthPx / intrinsicWidth;
			}

			var userUnits = !!(options && options.userUnits);
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
				if (!stroke || stroke === 'none')
					continue;
				// non-scaling strokes are already fixed (or deliberate); the PDF writer ignores vector-effect, see below
				var vectorEffect = element.getAttribute('vector-effect');
				if (vectorEffect && (!userUnits || vectorEffect !== 'non-scaling-stroke'))
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
				var effectiveScale = AscCommon.EoSvgEffectiveScale(transforms, factor);
				if (userUnits)
				{
					// for renderers without vector-effect (the PDF writer): the width in the units of the element.
					// A non-scaling stroke (the picture was already normalised for the screen) is a width in pixels:
					// a hairline gets the PDF hairline width, a wider one keeps its pixel width.
					var widthPx = vectorEffect ? (strokeWidth < 1 ? minPx : strokeWidth) : strokeWidth * effectiveScale * destScale;
					if (!vectorEffect && widthPx >= minPx)
						continue;
					if (vectorEffect)
						element.removeAttribute('vector-effect');
					element.setAttribute('stroke-width', String(parseFloat((Math.max(widthPx, minPx) / (effectiveScale * destScale)).toPrecision(6))));
				}
				else if (strokeWidth * effectiveScale * destScale >= minPx)
				{
					continue;
				}
				else
				{
					element.setAttribute('vector-effect', 'non-scaling-stroke');
					element.setAttribute('stroke-width', String(Math.round(minPx * 1000) / 1000));
				}
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
	AscCommon['EoSvgHairlinePx'] = AscCommon.EoSvgHairlinePx;
	AscCommon['EoSvgHairlinePxForDpi'] = AscCommon.EoSvgHairlinePxForDpi;
	AscCommon['EoSvgEffectiveScale'] = AscCommon.EoSvgEffectiveScale;
})(window);
