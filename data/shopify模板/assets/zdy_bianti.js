/*获取指定id的dom */
let zdy_product_form = document.getElementById("product-form-" + zdy_product_id);
let zdy_product_variant_input = zdy_product_form.querySelector('input[name="id"]');

let zdy_styleTag = document.getElementById('zdy-bianti-style');


if (zdy_product_variant_input) {
    let zdy_variant_id = zdy_product_variant_input.value;
    let zdy_variant_obj = zdy_product_bianti.find(e => e.id == zdy_variant_id);
    let zdy_css_content = `.zdy-bianti-${zdy_product_id} > section[data-sku="${zdy_variant_obj.sku}"]{ display: block; }`;
    zdy_styleTag.innerHTML = zdy_css_content;
}

zdy_product_variant_input.addEventListener('change', function () {
    let zdy_variant_id = zdy_product_variant_input.value;
    let zdy_variant_obj = zdy_product_bianti.find(e => e.id == zdy_variant_id);
    let zdy_css_content = `.zdy-bianti-${zdy_product_id} > section[data-sku="${zdy_variant_obj.sku}"]{ display: block; }`;
    zdy_styleTag.innerHTML = zdy_css_content;
});
