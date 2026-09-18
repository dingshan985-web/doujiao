/**
 * Product subtotal under quantity selector
 * - Subtotal = unit price x qty, updates on qty change and variant change
 * - Zero dependencies on theme.js internals (safe if theme.js partially fails)
 */
(function () {
  "use strict";

  function formatMoney(cents) {
    try {
      if (window.Shopify && Shopify.formatMoney) {
        var fmt =
          (window.themeGlobalVariables &&
            themeGlobalVariables.settings &&
            themeGlobalVariables.settings.money_format) ||
          (window.cartStrings && cartStrings.money_format) ||
          "${{amount}}";
        return Shopify.formatMoney(cents, fmt);
      }
    } catch (e) {}
    return "$" + (cents / 100).toFixed(2);
  }

  function initScope(root) {
    var wraps = (root || document).querySelectorAll("[data-product-subtotal]");
    wraps.forEach(function (wrap) {
      if (wrap.dataset.subtotalInit === "1") return;
      wrap.dataset.subtotalInit = "1";

      var priceEl = wrap.querySelector("[data-subtotal-price]");
      if (!priceEl) return;

      var container = wrap.closest("product-form") || wrap.closest("form") || document;
      var qtyInput = container.querySelector('input[name="quantity"]');
      var idInput = container.querySelector("input.product-variant-id") || container.querySelector('input[name="id"]');
      if (!qtyInput) return;

      var basePrice = parseInt(priceEl.getAttribute("data-variant-price"), 10) || 0;

      // variants JSON (rendered by variant picker) to resolve price on variant change
      var variantsList = null;
      try {
        var scope = wrap.closest(".product__item-js") || wrap.closest(".shopify-section") || document;
        var jsonEl = scope.querySelector('.productVariantsQty, script[type="application/json"]');
        // prefer a JSON whose entries look like variants (have id + price)
        document.querySelectorAll("script[type='application/json']").forEach(function (el) {
          if (variantsList) return;
          try {
            var data = JSON.parse(el.textContent);
            if (Array.isArray(data) && data.length && data[0] && "price" in data[0] && "id" in data[0]) {
              // only accept one inside the same product section
              if (el.closest(".shopify-section") === (wrap.closest(".shopify-section") || document)) {
                variantsList = data;
              }
            }
          } catch (err) {}
        });
      } catch (e) {}

      function update() {
        var qty = parseInt(qtyInput.value, 10);
        if (isNaN(qty) || qty < 1) qty = 1;
        priceEl.textContent = formatMoney(basePrice * qty);
      }

      // qty changes: + / - buttons and direct typing
      container.addEventListener("click", function (e) {
        var btn = e.target.closest(".quantity__button");
        if (!btn) return;
        setTimeout(update, 0);
        setTimeout(update, 60);
      });
      qtyInput.addEventListener("change", update);
      qtyInput.addEventListener("input", update);
      qtyInput.addEventListener("keyup", update);

      // variant change: theme rewrites the hidden variant id input value
      if (idInput && "MutationObserver" in window) {
        new MutationObserver(function () {
          if (variantsList) {
            for (var i = 0; i < variantsList.length; i++) {
              if (String(variantsList[i].id) === String(idInput.value)) {
                basePrice = variantsList[i].price;
                break;
              }
            }
          }
          update();
        }).observe(idInput, { attributes: true, attributeFilter: ["value"] });
      }

      update();
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      initScope(document);
    });
  } else {
    initScope(document);
  }
  // theme editor / section re-render support
  document.addEventListener("shopify:section:load", function (e) {
    initScope(e.target);
  });
})();
