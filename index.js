/**
 * Eco-growth Discovery Project
 * Client-side script for the Office Add-in.
 */

async function getAuthToken() {
  try {
    if (typeof Office !== "undefined" && Office.auth && typeof Office.auth.getAccessToken === "function") {
      return await Office.auth.getAccessToken();
    }

    if (typeof console !== "undefined" && console.warn) {
      console.warn("Office.auth.getAccessToken is not available in this environment.");
    }
    return null;
  } catch (error) {
    if (typeof console !== "undefined" && console.error) {
      console.error("Error getting access token:", error);
    }
    return null;
  }
}

async function fetchSources(token) {
  try {
    const headers = {
      "Content-Type": "application/json"
    };

    if (token) {
      headers.Authorization = `Bearer ${token}`;
    }

    const response = await fetch("https://jules.googleapis.com/v1alpha/sources", {
      headers
    });

    if (!response.ok) {
      throw new Error(`API request failed with status ${response.status}`);
    }

    return await response.json();
  } catch (error) {
    if (typeof console !== "undefined" && console.error) {
      console.error("Error fetching sources:", error);
    }
    return null;
  }
}

function renderSources(data) {
  if (typeof document === "undefined") {
    return [];
  }

  const apiResponseDisplay = document.getElementById("api-response-display");
  if (!apiResponseDisplay) {
    return [];
  }

  const sources = Array.isArray(data && data.sources) ? data.sources : [];

  if (!sources.length) {
    apiResponseDisplay.textContent = "No sources found.";
    return [];
  }

  apiResponseDisplay.innerHTML = "";
  sources.forEach((source, index) => {
    const item = document.createElement("div");
    item.className = "source-item";

    const displayName = typeof source?.name === "string" ? source.name : `Source ${index + 1}`;
    item.textContent = displayName;
    apiResponseDisplay.appendChild(item);
  });

  return sources;
}

function initializeOfficeAddIn(info) {
  if (!info || !info.host) {
    return;
  }

  if (info.host === Office.HostType.PowerPoint || info.host === "PowerPoint") {
    const initialsDisplay = document.getElementById("initials-display");
    if (initialsDisplay) {
      initialsDisplay.textContent = "JV";
    }

    const apiResponseDisplay = document.getElementById("api-response-display");
    if (apiResponseDisplay) {
      apiResponseDisplay.textContent = "Loading sources...";
    }

    getAuthToken().then(async (token) => {
      if (!token) {
        if (apiResponseDisplay) {
          apiResponseDisplay.textContent = "Authentication unavailable.";
        }
        return;
      }

      const data = await fetchSources(token);
      renderSources(data);
    });
  }
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    if (typeof console !== "undefined" && console.log) {
      console.log("Eco-growth Discovery initialized.");
    }

    initializeOfficeAddIn(info);
  });
} else if (typeof window !== "undefined") {
  window.addEventListener("DOMContentLoaded", () => {
    const initialsDisplay = document.getElementById("initials-display");
    if (initialsDisplay) {
      initialsDisplay.textContent = "JV";
    }
  });
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { getAuthToken, fetchSources, renderSources };
}
