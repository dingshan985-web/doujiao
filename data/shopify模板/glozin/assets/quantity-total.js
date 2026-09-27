(() => {
  const totalSelector = ".product-quantity-total";

  const getMoneyFormat = () =>
    window.cartStrings?.money_format ||
    window.themeGlobalVariables?.settings?.money_format ||
    window.money_format?.amount ||
    "${{amount}}";

  const formatMoney = (cents) => {
    if (window.Shopify?.formatMoney) {
      return Shopify.formatMoney(Math.round(cents), getMoneyFormat());
    }
    return (Number(cents || 0) / 100).toFixed(2);
  };

  const getScope = (total) =>
    total.closest(".product__item-js") || total.closest("section") || document;

  const getForm = (total, scope) =>
    total.closest("product-form") ||
    total.closest("form") ||
    scope.querySelector('form[data-type="add-to-cart-form"]');

  const getVariantPrice = (total, scope, form) => {
    const variantInput = form?.querySelector(
      "input.product-variant-id, input[name='id']"
    );
    const variantId = variantInput?.value;
    const productJson = scope.querySelector("[data-product-json], .productJson");

    if (variantId && productJson) {
      try {
        const product = JSON.parse(productJson.textContent);
        const variant = product.variants?.find(
          (item) => String(item.id) === String(variantId)
        );
        if (variant?.price != null) return Number(variant.price);
      } catch (error) {
        // Fall back to the Liquid-rendered initial price if the JSON is unavailable.
      }
    }

    return Number(total.dataset.unitPrice || 0);
  };

  const updateTotal = (total) => {
    const scope = getScope(total);
    const form = getForm(total, scope);
    const quantityInput =
      form?.querySelector('input[name="quantity"]') ||
      scope.querySelector('input[name="quantity"]');
    const quantity = Math.max(0, Number(quantityInput?.value || 0));
    const unitPrice = getVariantPrice(total, scope, form);
    const valueElement = total.querySelector(".product-quantity-total__value");

    if (!valueElement) return;
    valueElement.innerHTML = formatMoney(unitPrice * quantity);
    total.dataset.unitPrice = String(unitPrice);
  };

  const init = (total) => {
    if (total.dataset.quantityTotalInitialized === "true") return;
    total.dataset.quantityTotalInitialized = "true";

    const scope = getScope(total);
    const form = getForm(total, scope);
    const quantityInput =
      form?.querySelector('input[name="quantity"]') ||
      scope.querySelector('input[name="quantity"]');

    if (quantityInput) {
      quantityInput.addEventListener("input", () => updateTotal(total));
      quantityInput.addEventListener("change", () => updateTotal(total));
    }

    form
      ?.querySelectorAll("input.product-variant-id, input[name='id']")
      .forEach((input) => {
        input.addEventListener("change", () => updateTotal(total));
      });

    updateTotal(total);
  };

  const initAll = () => document.querySelectorAll(totalSelector).forEach(init);

  const updateScopeTotals = (target) => {
    const scope =
      target.closest(".product__item-js") || target.closest("section") || document;
    scope.querySelectorAll(totalSelector).forEach(updateTotal);
  };

  // Use delegated listeners as a fallback for quantity inputs that are
  // replaced by Shopify section rendering or by a quick-view form.
  document.addEventListener("input", (event) => {
    if (event.target.matches('input[name="quantity"]')) {
      updateScopeTotals(event.target);
    }
  });
  document.addEventListener("change", (event) => {
    if (
      event.target.matches('input[name="quantity"]') ||
      event.target.matches("input.product-variant-id, input[name='id']")
    ) {
      updateScopeTotals(event.target);
    }
  });

  initAll();
  document.addEventListener("DOMContentLoaded", initAll, { once: true });

  new MutationObserver(initAll).observe(document.documentElement, {
    childList: true,
    subtree: true,
  });
})();
