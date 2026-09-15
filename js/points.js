const Points = {
  rewards: {
    Sold: 100,
    "Do Not Call": 3,
    FNQ: 3,
    "Pending Order": 3,
    "Already has Fiber": 5,
    TDM: 10,
    "1st Contact": 5,
    "2nd Contact": 10,
    "3rd Contact": 15,
    callbackSet: 10,
    dailyLogin: 5,
  },

  wakeUpHUD() {
    const hud = document.getElementById("economy-hud");
    if (!hud) return;

    // Slide it up!
    hud.classList.remove("hud-hidden");

    // Clear any existing sleep countdowns
    clearTimeout(this._hudSleepTimer);

    // Set a new timer to hide it after 5 seconds of inactivity
    this._hudSleepTimer = setTimeout(() => {
      this._hudSleepTimer = null; // Timer is officially done

      // Only hide it if the user isn't currently hovering over it!
      if (!this._isMouseNearBottom) {
        hud.classList.add("hud-hidden");
      }
    }, 5000);
  },

  initHUDAutoHider() {
    // Wake it up on initial load so they see their balance when they log in
    this.wakeUpHUD();

    // 🧠 THE BRAIN: We abstract the logic so both Mouse and Touch can use it
    const checkPointerPosition = (clientY) => {
      // If the pointer (mouse or finger) is within the bottom 100 pixels
      if (clientY > window.innerHeight - 100) {
        this._isMouseNearBottom = true;
        const hud = document.getElementById("economy-hud");
        if (hud) hud.classList.remove("hud-hidden");
      } else {
        this._isMouseNearBottom = false;

        // If they move away, AND the 5-second sleep timer isn't running...
        if (!this._hudSleepTimer) {
          const hud = document.getElementById("economy-hud");
          if (hud) hud.classList.add("hud-hidden");
        }
      }
    };

    // 🖱️ DESKTOP: Track the mouse globally
    document.addEventListener("mousemove", (e) => {
      checkPointerPosition(e.clientY);
    });

    // 📱 MOBILE: Track when a finger first taps the screen
    document.addEventListener(
      "touchstart",
      (e) => {
        // e.touches[0] gets the coordinates of the first finger on the glass
        checkPointerPosition(e.touches[0].clientY);
      },
      { passive: true },
    );

    // 📱 MOBILE: Track when a finger drags across the screen (swiping)
    document.addEventListener(
      "touchmove",
      (e) => {
        checkPointerPosition(e.touches[0].clientY);
      },
      { passive: true },
    );
  },

  async fetchBalances() {
    console.log("🏦 Fetching agent point balances...");

    try {
      // 1. Ask graph.js to go get the raw data
      const allScores = await Graph.getAgentScores();
      State.agentScores = allScores;

      // 2. Figure out who is currently logged in
      const userEmail = ((State.currentUser && State.currentUser.email) || "")
        .toLowerCase()
        .trim();
      const userName =
        (State.currentUser && State.currentUser.name) || userEmail;

      // 3. Find the current user's specific row
      let myScore = allScores.find(
        (s) => (s.AgentEmail || "").toLowerCase().trim() === userEmail,
      );

      // 4. THE INTERCEPT: If they don't exist in the database, build their account!
      if (!myScore && userEmail) {
        console.log(
          `🆕 No point bank found for ${userName}. Creating a new account...`,
        );
        myScore = await Graph.createAgentScore(userEmail, userName);

        // Push this brand new score into the State array so the leaderboard sees them
        State.agentScores.push(myScore);
      }

      // 5. Update our local bank (Using optional chaining fallback just in case)
      this.currentBalance = myScore ? myScore.CurrentPoints : 0;
      this.lifetimePoints = myScore ? myScore.LifetimePoints : 0;

      console.log(
        `💰 Bank Loaded! Balance: ${this.currentBalance} | Lifetime: ${this.lifetimePoints}`,
      );
    } catch (error) {
      console.error("Failed to load point balances:", error);
    }
  },

  getDailySalesCount() {
    const myName = (State.currentUser && State.currentUser.name) || "";
    if (!myName) return 0;

    const uniqueSalesToday = new Set();
    const todayString = new Date().toDateString(); // ⚡ Establish Today

    // 🧽 THE FIX: Give the HUD the Universal Name Cleaner too!
    const cleanName = (str) =>
      (str || "")
        .replace(/[^\w\s]/gi, "")
        .toLowerCase()
        .trim();

    (State.activityLog || []).forEach((log) => {
      // ⚡ Check the date!
      const logDate = new Date(log.timestamp).toDateString();

      if (
        log.action === "Status: " + Config.soldStatus &&
        logDate === todayString
      ) {
        const pastLead = (State.leads || []).find(
          (l) => String(l.id) === String(log.leadId),
        );

        if (pastLead) {
          // ⚡ Compare the scrubbed names!
          const isMine = cleanName(pastLead.assignedTo) === cleanName(myName);
          const isStillSold = pastLead.status === Config.soldStatus;

          if (isMine && isStillSold) {
            uniqueSalesToday.add(log.leadId);
          }
        }
      }
    });

    return uniqueSalesToday.size;
  },

  async awardPoints(actionType, leadId = null) {
    let pointValue = this.rewards[actionType];
    if (!pointValue) return false;

    let beneficiaryEmail = (State.currentUser?.email || "")
      .toLowerCase()
      .trim();
    let beneficiaryName = State.currentUser?.name || "";

    if (Config.leadStatuses?.includes(actionType) && leadId) {
      const realLead = (State.leads || []).find(
        (l) => String(l.id) === String(leadId),
      );
      if (!realLead || realLead.status !== actionType) return false;

      if (realLead.assignedTo) {
        const normalize = (str) =>
          (str || "")
            .replace(/[^\w\s]/gi, "")
            .replace(/\s+/g, " ")
            .toLowerCase()
            .trim();
        const normalizedLeadOwner = normalize(realLead.assignedTo);
        const ownerBankRow = (State.agentScores || []).find(
          (s) => normalize(s.AgentName) === normalizedLeadOwner,
        );

        if (ownerBankRow?.AgentEmail) {
          beneficiaryEmail = ownerBankRow.AgentEmail.toLowerCase().trim();
          beneficiaryName = ownerBankRow.AgentName;
        } else {
          return false;
        }
      }
    }

    if (!beneficiaryEmail) return false;

    const uniqueSalesToday = new Set();
    const todayString = new Date().toDateString();
    const cleanName = (str) =>
      (str || "")
        .replace(/[^\w\s]/gi, "")
        .toLowerCase()
        .trim();

    (State.activityLog || []).forEach((log) => {
      if (
        log.action === "Status: " + Config.soldStatus &&
        log.leadId !== leadId &&
        new Date(log.timestamp).toDateString() === todayString
      ) {
        const pastLead = (State.leads || []).find(
          (l) => String(l.id) === String(log.leadId),
        );
        if (
          pastLead &&
          cleanName(pastLead.assignedTo) === cleanName(beneficiaryName) &&
          pastLead.status === Config.soldStatus
        ) {
          uniqueSalesToday.add(log.leadId);
        }
      }
    });

    const previousSalesCount = uniqueSalesToday.size;
    if (actionType === Config.soldStatus) {
      pointValue *= previousSalesCount + 1;
    } else if (previousSalesCount > 0) {
      pointValue += previousSalesCount * 2;
    }

    try {
      // 🚀 THE PATCH: Passing the dynamically resolved beneficiaryEmail
      if (
        leadId &&
        (await Graph.checkLedgerForDuplicate(
          leadId,
          actionType,
          beneficiaryEmail,
        ))
      )
        return false;

      const myScoreIndex = State.agentScores.findIndex(
        (s) => s.AgentEmail.toLowerCase() === beneficiaryEmail,
      );
      let myScoreId = null,
        newCurrentPoints = pointValue,
        newLifetimePoints = pointValue;

      if (myScoreIndex !== -1) {
        newCurrentPoints = State.agentScores[myScoreIndex].CurrentPoints +=
          pointValue;
        newLifetimePoints = State.agentScores[myScoreIndex].LifetimePoints +=
          pointValue;
        myScoreId = State.agentScores[myScoreIndex].id;
      }

      const currentUserEmail = (State.currentUser?.email || "")
        .toLowerCase()
        .trim();
      if (beneficiaryEmail === currentUserEmail) {
        const oldLevel = this.calculateLevel(this.lifetimePoints);
        const newLevel = this.calculateLevel(newLifetimePoints);
        this.currentBalance = newCurrentPoints;
        this.lifetimePoints = newLifetimePoints;

        let maxP =
          pointValue >= 400
            ? 50
            : pointValue >= 300
              ? 40
              : pointValue >= 200
                ? 35
                : pointValue >= 100
                  ? 25
                  : 12;
        const particleCount =
          Math.floor(Math.random() * (maxP - Math.ceil(maxP * 0.25) + 1)) +
          Math.ceil(maxP * 0.25);
        this.flyParticlesToHUD(particleCount, pointValue);

        if (newLevel > oldLevel) Points.triggerLevelUp(newLevel);
      }

      await Graph.writeLedgerTransaction(
        beneficiaryEmail,
        actionType,
        pointValue,
        leadId,
      );
      if (myScoreId)
        await Graph.updateAgentScore(
          myScoreId,
          newCurrentPoints,
          newLifetimePoints,
        );
      return true;
    } catch (error) {
      return false;
    }
  },

  async purchaseItem(itemId, cost) {
    // Subtract points and log the purchase
  },

  // ==========================================
  // 📈 THE PROGRESSION SYSTEM (XP & LEVELS) 📈
  // ==========================================

  // Calculates the agent's current level based on their Lifetime Points (XP)
  calculateLevel(xp) {
    // Square Root Curve: Divisor of 150. Level 50 takes ~360k XP.
    return Math.floor(Math.sqrt((xp || 0) / 125)) + 1;
  },

  // Calculates exactly how much total XP is required to reach a specific level
  xpForNextLevel(targetLevel) {
    // To find out how much XP is needed to reach Level X
    return Math.pow(targetLevel - 1, 2) * 125;
  },

  triggerLevelUp(newLevel) {
    const banner = document.getElementById("level-up-banner");
    const numSpan = document.getElementById("lvl-up-number");

    if (!banner || !numSpan) return;

    // Update the number
    numSpan.textContent = newLevel;

    // Slam it onto the screen
    banner.classList.add("is-active");

    // Trigger an absolute MAX OVERLOAD of particles (100 is safe for the HD 530!)
    // They will explode out right as the banner slams in.
    this.flyParticlesToHUD(50);

    // Leave it on screen for 3.5 seconds, then fade it out
    setTimeout(() => {
      banner.classList.remove("is-active");
    }, 3500);
  },

  updateHUD(glowDuration = 800) {
    const elLevel = document.getElementById("hud-level");
    const elFill = document.getElementById("hud-fill");
    const elXpCurrent = document.getElementById("hud-xp-current");
    const elXpNext = document.getElementById("hud-xp-next");
    const elBalance = document.getElementById("hud-balance");

    if (!elLevel || !elFill) return; // Fail silently if HUD isn't on the page

    // 1. Calculate the progression stats
    const currentLevel = this.calculateLevel(this.lifetimePoints);
    const xpFloor = this.xpForNextLevel(currentLevel); // XP required for CURRENT level
    const xpCeiling = this.xpForNextLevel(currentLevel + 1); // XP required for NEXT level

    // 2. Math for the Progress Bar
    const xpEarnedInCurrentTier = this.lifetimePoints - xpFloor;
    const xpNeededForNextTier = xpCeiling - xpFloor;
    let fillPercentage = (xpEarnedInCurrentTier / xpNeededForNextTier) * 100;

    // Prevent visual overflow
    if (fillPercentage > 100) fillPercentage = 100;
    if (fillPercentage < 0) fillPercentage = 0;

    // 3. Update the DOM
    const currentDOMLevel = parseInt(elLevel.textContent) || 1;

    elLevel.textContent = currentLevel;
    const oldWidth = elFill.style.width;

    // ⚡ THE SNAP TO ZERO FIX ⚡
    // If the new level is higher than what was just on the screen...
    if (currentLevel > currentDOMLevel) {
      // 1. Turn off the CSS animation completely
      elFill.style.transition = "none";
      // 2. Snap the bar to 0%
      elFill.style.width = "0%";
      // 3. THE MAGIC TRICK: Asking for offsetWidth forces the browser to redraw the screen immediately!
      void elFill.offsetWidth;
      // 4. Turn the CSS animation back on (empty string resets it to your stylesheet rules)
      elFill.style.transition = "";
    }

    // Now set the actual new destination, and it will slide smoothly from 0!
    elFill.style.width = `${fillPercentage}%`;

    // 🌟 TRIGGER THE BAR GLOW 🌟
    if (oldWidth && oldWidth !== elFill.style.width) {
      elFill.classList.add("is-filling");
      clearTimeout(this._glowTimer);

      this._glowTimer = setTimeout(() => {
        if (elFill) elFill.classList.remove("is-filling");
      }, glowDuration);
    }

    // Show exact numbers: e.g., "45 / 375 XP"
    elXpCurrent.textContent = xpEarnedInCurrentTier.toLocaleString();
    elXpNext.textContent = xpNeededForNextTier.toLocaleString();

    // Update the spendable bank
    elBalance.textContent = this.currentBalance.toLocaleString();

    const elComboContainer = document.getElementById("hud-combo-container");
    const elComboCount = document.getElementById("hud-combo-count");

    if (elComboContainer && elComboCount) {
      const currentSales = this.getDailySalesCount();
      const oldSales = parseInt(elComboCount.textContent) || 0;

      // 1. Update the number
      elComboCount.textContent = currentSales;

      // 2. Light it up if they have at least 1 sale
      if (currentSales >= 3) {
        // 🔥 SUPER STREAK: 3+ Sales
        elComboContainer.classList.add("is-hot", "on-fire");
      } else if (currentSales > 0) {
        // ⚡ REGULAR COMBO: 1-2 Sales
        elComboContainer.classList.add("is-hot");
        elComboContainer.classList.remove("on-fire");
      } else {
        // 💤 SLEEPING: 0 Sales
        elComboContainer.classList.remove("is-hot", "on-fire");
      }

      // 3. Trigger the violent bump animation if the number went UP!
      if (currentSales > oldSales) {
        // The Forced Reflow trick to restart the animation instantly
        elComboCount.classList.remove("combo-bump");
        void elComboCount.offsetWidth;
        elComboCount.classList.add("combo-bump");
      }
    }
  },

  // ==========================================
  // ✨ THE DOPAMINE PARTICLE ENGINE (v2.2) ✨
  // ==========================================
  flyParticlesToHUD(particleCount = 5, pointsEarned = 0) {
    const hud = document.getElementById("economy-hud");
    const wasHidden = hud && hud.classList.contains("hud-hidden");
    this.wakeUpHUD();
    const mathDelay = wasHidden ? 350 : 0;

    // Delay the math and the explosion until the HUD is securely locked in place
    setTimeout(() => {
      const targetEl = document.querySelector(".hud-badge");
      if (!targetEl) {
        this.updateHUD(); // Fallback if HUD is broken
        return;
      }

      // 1. Set the Origin (Now guaranteed to be accurate!)
      const targetRect = targetEl.getBoundingClientRect();
      const originX = targetRect.left + targetRect.width / 2;
      const originY = targetRect.top + targetRect.height / 2;

      let maxTotalTime = 0;
      let minTotalTime = Infinity;

      for (let i = 0; i < particleCount; i++) {
        // ⚡ THE FIX: Stagger the spawn by 8 milliseconds per particle
        const spawnDelay = i * 5;

        // --- RANDOMIZED TIMINGS FOR STAGGERING ---
        // (Calculated synchronously so our total times update correctly!)
        const burstDuration = 300 + Math.random() * 300;
        const restDuration = 200 + Math.random() * 500;
        const returnDuration = 400 + Math.random() * 500;

        const totalTime =
          spawnDelay + burstDuration + restDuration + returnDuration;
        if (totalTime > maxTotalTime) maxTotalTime = totalTime;
        if (totalTime < minTotalTime) minTotalTime = totalTime;

        // Delay the actual DOM creation and animation
        setTimeout(() => {
          const particle = document.createElement("div");
          particle.className = "xp-particle";
          document.body.appendChild(particle);

          // Randomize Size (between 6px and 16px)
          const size = Math.random() * 10 + 6;
          particle.style.width = `${size}px`;
          particle.style.height = `${size}px`;

          // Spawn exactly at the center of the HUD badge
          particle.style.left = `${originX - size / 2}px`;
          particle.style.top = `${originY - size / 2}px`;

          particle.style.opacity = "0";
          particle.style.transform = "translate(0, 0) scale(0)";

          // Desynchronize the strobe
          particle.style.animationDelay = `-${Math.random()}s`;

          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              // PHASE 1: THE BURST
              const minAngle = 205 * (Math.PI / 180);
              const angleRange = 130 * (Math.PI / 180);

              const angle = minAngle + Math.random() * angleRange;

              const distance = 80 + Math.random() * 160;

              const burstX = Math.cos(angle) * distance;
              const burstY = Math.sin(angle) * distance;

              particle.style.transition = `transform ${burstDuration}ms cubic-bezier(0.175, 0.885, 0.32, 1.275), opacity 0.2s ease-out`;
              particle.style.opacity = "1";
              particle.style.transform = `translate(${burstX}px, ${burstY}px) scale(1)`;

              // PHASE 2 & 3: THE REST & RETURN
              setTimeout(() => {
                particle.style.transition = `transform ${returnDuration}ms cubic-bezier(0.55, 0.085, 0.68, 0.53), opacity ${returnDuration}ms ease-in`;
                particle.style.transform = "translate(0px, 0px) scale(0.3)";
                particle.style.opacity = "0.5";

                // PHASE 4: THE IMPACT
                setTimeout(() => {
                  particle.remove();
                  targetEl.style.transform = "scale(1.1)";
                  setTimeout(() => (targetEl.style.transform = "scale(1)"), 50);
                }, returnDuration);
              }, burstDuration + restDuration);
            });
          });
        }, spawnDelay); // <--- Matches the offset for this specific particle!
      }

      // ⚡ PHASE 5: THE BANK DEPOSIT TRIGGER ⚡
      setTimeout(() => {
        const barrageDuration = maxTotalTime - minTotalTime;
        this.updateHUD(Math.max(800, barrageDuration));

        // 💬 FIRE THE FLOATING TEXT!
        this.spawnFloatingText(pointsEarned);
      }, minTotalTime);
    }, mathDelay); // <--- This holds everything back until the HUD stops moving!
  },

  spawnFloatingText(pointsEarned) {
    if (!pointsEarned) return;

    const badge = document.querySelector(".hud-badge");
    if (!badge) return;

    const rect = badge.getBoundingClientRect();
    const text = document.createElement("div");
    text.className = "floating-xp-text";
    text.textContent = `+${pointsEarned}`;

    // Position it slightly above and to the right of the Level Badge
    text.style.left = `${rect.left + 40}px`;
    text.style.top = `${rect.top - 10}px`;

    document.body.appendChild(text);

    // Clean it up after the 1.2s CSS animation finishes
    setTimeout(() => {
      text.remove();
    }, 1200);
  },
};
