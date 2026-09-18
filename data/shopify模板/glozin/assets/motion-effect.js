class MotionElement extends HTMLElement {
  constructor() {
    super();
  }
  connectedCallback() {
    if (this.isHold) return;
    this.preInitialize();
    if (this.dataset.noview) {
      return;
    }
    if (!this.isConnected || !document.body.contains(this)) {
      setTimeout(() => {
        this.connectedCallback();
      }, 100);
      return;
    }

    var _this = this;
    let custom_margin = this.dataset.margin || "0px 0px -5px 0px";
    let rootElement = {};
    if (this.dataset.parent) {
      var parent = this.closest(`.${this.dataset.parent}`);
      rootElement = { root: parent };
      custom_margin = "0px 0px 0px 0px";
    }
    Motion.inView(
      _this,
      async () => {
        if (
          !this.isInstant &&
          this.mediaElements &&
          this.hasAttribute("data-image")
        ) {
          await imageReady(this.mediaElements);
        }
        setTimeout(() => {
          _this.initialize();
        }, 50);
      },
      { margin: custom_margin, ...rootElement }
    );
  }

  get isHold() {
    return this.hasAttribute("hold");
  }

  get slideElement() {
    return this.hasAttribute("slideshow") ? true : false;
  }

  get Transition() {
    let transition = this.getAttribute("data-transition")
      ?.split(",")
      .map(Number);
    return transition || [0.22, 0.61, 0.36, 1];
  }

  get isInstant() {
    return this.hasAttribute("data-instantly");
  }

  get mediaElements() {
    return Array.from(this.querySelectorAll("img, iframe, svg"));
  }

  get animationType() {
    return this.dataset.motion || "fade-up";
  }

  get animationDelay() {
    return parseInt(this.dataset.motionDelay || 0) / 1000;
  }

  set animationDelay(value) {
    this.dataset.motionDelay = value;
  }

  get delayLoad() {
    return this.classList.contains("animate-delay");
  }

  preInitialize() {
    if (window.innerWidth < 768 && !this.slideElement) return;
    if (this.isHold) return;
    switch (this.animationType) {
      case "fade-in":
        Motion.animate(this, { opacity: 0 }, { duration: 0 });
        break;

      case "fade-up":
        Motion.animate(
          this,
          { transform: "translateY(2.5rem)", opacity: 0 },
          { duration: 0 }
        );
        break;

      case "fade-up-sm":
        Motion.animate(
          this,
          { transform: "translateY(1rem)", opacity: 0 },
          { duration: 0 }
        );
        break;

      case "fade-up-lg":
        Motion.animate(
          this,
          { transform: "translateY(3rem)", opacity: 0 },
          { duration: 0 }
        );
        break;

      case "zoom-in":
        Motion.animate(this, { transform: "scale(0.8)" }, { duration: 0 });
        break;
      case "zoom-in-lg":
        Motion.animate(this, { transform: "scale(0)" }, { duration: 0 });
        break;

      case "zoom-out":
        Motion.animate(this, { transform: "scale(1.07)" }, { duration: 0 });
        break;

      case "zoom-out-sm":
        Motion.animate(this, { transform: "scale(1.07)" }, { duration: 0 });
        break;
      case "zoom-out-lg":
        Motion.animate(this, { transform: "scale(1.07)" }, { duration: 0 });
        break;
    }
  }

  async initialize() {
    if (window.innerWidth < 768 && !this.slideElement) return;
    if (this.delayLoad) return;
    if (this.isHold) return;
    switch (this.animationType) {
      case "fade-in":
        await Motion.animate(
          this,
          { opacity: 1 },
          { duration: 1.5, delay: this.animationDelay, easing: this.transition }
        ).finished;
        break;

      case "fade-up":
        await Motion.animate(
          this,
          { transform: "translateY(0)", opacity: 1 },
          {
            duration: 0.5,
            delay: this.animationDelay,
            easing: this.transition,
          }
        ).finished;
        break;

      case "fade-up-sm":
        await Motion.animate(
          this,
          { transform: "translateY(0)", opacity: 1 },
          {
            duration: 0.3,
            delay: this.animationDelay,
            easing: this.transition,
          }
        ).finished;
        break;

      case "fade-up-lg":
        await Motion.animate(
          this,
          { transform: "translateY(0)", opacity: 1 },
          {
            duration: 0.5,
            delay: this.animationDelay,
            easing: this.transition,
          }
        ).finished;
        break;

      case "zoom-in":
        await Motion.animate(
          this,
          { transform: "scale(1)" },
          { duration: 1.3, delay: this.animationDelay, easing: this.transition }
        ).finished;
        break;

      case "zoom-in-lg":
        await Motion.animate(
          this,
          { transform: "scale(1)" },
          { duration: 0.5, delay: this.animationDelay, easing: this.transition }
        ).finished;
        break;

      case "zoom-out":
        await Motion.animate(
          this,
          { transform: "scale(1)" },
          { duration: 1.5, delay: this.animationDelay, easing: this.transition }
        ).finished;
        break;

      case "zoom-out-sm":
        await Motion.animate(
          this,
          { transform: "scale(1)" },
          { duration: 1, delay: this.animationDelay, easing: this.transition }
        ).finished;
        break;
      case "zoom-out-lg":
        await Motion.animate(
          this,
          { transform: "scale(1)" },
          { duration: 1, delay: this.animationDelay, easing: this.transition }
        ).finished;
        break;
    }
    if (this.slideElement) {
      this.classList.add("animate-delay");
    }
  }
  refreshAnimation() {
    this.removeAttribute("hold");
    this.preInitialize();
    setTimeout(() => {
      this.initialize();
    }, 50); // Delay a bit to make animation re init properly.
  }
}
customElements.define("motion-element", MotionElement);

class MotionItemsEffect extends HTMLElement {
  constructor() {
    super();
    this.init();
  }

  init() {
    this.setupInitialAnimation();
    this.setupInViewEffect();
  }

  get allItems() {
    return this.querySelectorAll(".motion-item");
  }

  get visibleItems() {
    return this.querySelectorAll(".product-item:not([style])");
  }

  get animationType() {
    return this.dataset.motion || "fade-up";
  }

  setupInitialAnimation() {
    if (window.innerWidth < 768) {
      return;
    }
    switch (this.animationType) {
      case "fade-in":
        this.fadeInInit();
        break;
      case "fade-up-lg":
        this.fadeUpLgInit();
        break;
      case "zoom-in":
        this.zoomInInit();
        break;
    }
  }

  fadeInInit() {
    Motion.animate(
      this.allItems,
      {
        opacity: 0.01,
      },
      {
        duration: 0,
      }
    );
  }

  zoomInInit() {
    Motion.animate(
      this.allItems,
      {
        opacity: 0,
        scale: 0.8,
      },
      {
        duration: 0,
      }
    );
  }

  fadeUpLgInit() {
    Motion.animate(
      this.allItems,
      {
        y: 20,
        opacity: 0.01,
        visibility: "hidden",
      },
      {
        duration: 0,
      }
    );
  }

  setupInViewEffect() {
    if (window.innerWidth < 768) {
      return;
    }
    Motion.inView(this, this.animateItems.bind(this), {
      margin: "0px 0px -15px 0px",
    });
  }

  animateItems() {
    switch (this.animationType) {
      case "fade-in":
        this.fadeIn();
        break;
      case "fade-up-lg":
        this.fadeUpLg();
        break;
      case "zoom-in":
        this.zoomIn();
        break;
    }
  }

  fadeIn() {
    Motion.animate(
      this.allItems,
      {
        opacity: [0.01, 1],
      },
      {
        duration: 0.6,
        easing: [0.22, 0.61, 0.36, 1],
      }
    ).finished;
  }

  fadeUpLg() {
    Motion.animate(
      this.allItems,
      {
        y: [20, 0],
        opacity: [0.01, 1],
        visibility: ["hidden", "visible"],
      },
      {
        duration: 0.6,
        delay: Motion.stagger(0.15),
        easing: [0.22, 0.61, 0.36, 1],
      }
    ).finished;
  }

  zoomIn() {
    Motion.animate(
      this.allItems,
      {
        opacity: [0, 1],
        scale: [0.8, 1],
      },
      {
        duration: 0.6,
        delay: Motion.stagger(0.15),
        easing: [0.22, 0.61, 0.36, 1],
      }
    ).finished;
  }

  reloadAnimationEffect() {
    switch (this.animationType) {
      case "fade-in":
        this.fadeInReload();
        break;
      case "fade-up-lg":
        this.fadeUpLgReload();
        break;
      case "zoom-in":
        this.zoomInReload();
        break;
    }
  }

  fadeInReload() {
    Motion.animate(
      this.visibleItems,
      {
        opacity: [0.01, 1],
      },
      {
        duration: 0.6,
        easing: [0.22, 0.61, 0.36, 1],
      }
    ).finished;
  }

  fadeUpLgReload() {
    Motion.animate(
      this.visibleItems,
      {
        y: [20, 0],
        opacity: [0.01, 1],
        visibility: ["hidden", "visible"],
      },
      {
        duration: 0.6,
        delay: Motion.stagger(0.15),
        easing: [0.22, 0.61, 0.36, 1],
      }
    ).finished;
  }

  zoomInReload() {
    Motion.animate(
      this.visibleItems,
      {
        opacity: [0, 1],
        scale: [0.8, 1],
      },
      {
        duration: 0.6,
        delay: Motion.stagger(0.15),
        easing: [0.22, 0.61, 0.36, 1],
      }
    ).finished;
  }
}
customElements.define("motion-items-effect", MotionItemsEffect);
