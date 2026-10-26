/*
 * Histogram Pro licence check through Microsoft's Licensing API (API 4.7+).
 * Docs: learn.microsoft.com/power-bi/developer/visuals/licensing-api
 *
 * - Any plan of THIS offer in state Active or Warning means Pro. The offer has one paid plan,
 *   so the Partner Center service ID doesn't need to be hard-coded.
 * - Free features never depend on this. When Pro can't be confirmed (unsupported environment
 *   such as Publish to web or REST export, Desktop offline or signed out, an old host, an
 *   error, or a slow answer) the visual shows the free view. It never blocks the whole visual.
 * - No own purchase UI: missing licences are shown with Power BI's "feature blocked" banner.
 */
import powerbi from "powerbi-visuals-api";

import IVisualHost = powerbi.extensibility.visual.IVisualHost;

export type LicenceState =
    | "pending"      // still asking Power BI
    | "pro"          // usable Histogram Pro licence
    | "free"         // no usable licence
    | "unsupported"  // environment can't check licences (Publish to web, embed, RS, REST export)
    | "unavailable"; // couldn't check (Desktop offline or signed out, service error, old host, timeout)

/** How long to wait for Power BI before showing the free view. */
const TIMEOUT_MS = 3000;

export class ProLicence {
    state: LicenceState = "pending";
    /** Resolves once the state is final (never rejects). */
    readonly ready: Promise<LicenceState>;
    private manager: powerbi.extensibility.IVisualLicenseManager | undefined;
    private blockedShown = false;

    constructor(host: IVisualHost) {
        this.manager = (host as { licenseManager?: powerbi.extensibility.IVisualLicenseManager }).licenseManager;
        this.ready = this.check();
    }

    private async check(): Promise<LicenceState> {
        const m = this.manager;
        if (!m || typeof m.getAvailableServicePlans !== "function") return (this.state = "unavailable");
        let timer: ReturnType<typeof setTimeout> | undefined;
        const timeout = new Promise<"timeout">(res => { timer = setTimeout(() => res("timeout"), TIMEOUT_MS); });
        try {
            const r = await Promise.race([Promise.resolve(m.getAvailableServicePlans()), timeout]);
            if (r === "timeout") return (this.state = "unavailable");
            if (r.isLicenseUnsupportedEnv) return (this.state = "unsupported");
            if (!r.isLicenseInfoAvailable) return (this.state = "unavailable");
            const usable = (r.plans || []).some(p =>
                p.state === powerbi.ServicePlanState.Active || p.state === powerbi.ServicePlanState.Warning);
            return (this.state = usable ? "pro" : "free");
        } catch {
            return (this.state = "unavailable");
        } finally {
            if (timer !== undefined) clearTimeout(timer);
        }
    }

    /**
     * Power BI's "feature blocked" banner (10 s), once per time a Pro feature is switched on.
     * Only applies where Power BI supports licensing; elsewhere it's a no-op.
     */
    notifyBlocked(tooltip: string): void {
        if (this.blockedShown || this.state !== "free" || !this.manager) return;
        this.blockedShown = true;
        Promise.resolve(this.manager.notifyFeatureBlocked(tooltip.slice(0, 500))).catch((): void => undefined);
    }

    /** Call when no Pro feature is in use, so the next use shows the banner again. */
    reset(): void {
        if (!this.blockedShown) return;
        this.blockedShown = false;
        if (this.manager) Promise.resolve(this.manager.clearLicenseNotification()).catch((): void => undefined);
    }
}
