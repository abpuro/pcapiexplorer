import type { AccountInfo } from "@azure/msal-browser";
import { useEffect, useMemo, useRef, useState } from "react";
import { trackClarityEvent } from "./clarity";

export type PartnerCenterApiClient = <T>(endpoint: string, options?: RequestInit) => Promise<T>;

type CollectionResponse<T> = {
  totalCount?: number;
  items?: T[];
};

type GrowthMarginProductCode = {
  productCode: string;
  termId?: string;
};

type RequiredProduct = {
  productId: string;
  skuId: string;
  termDuration: string;
  billingCycle: string;
};

type PricingBenefit = {
  type?: string;
  value?: number;
};

type PricingPolicy = {
  policyId?: string;
  policyData?: {
    benefits?: PricingBenefit[];
  };
};

type OfferPolicy = {
  policyId?: string;
  policyData?: {
    beneficiary?: string;
    properties?: {
      startDate?: string;
      endDate?: string;
    };
    isOptional?: boolean;
  };
};

type ConstraintProduct = {
  BigId?: string;
  MinSeats?: number | null;
  MaxSeats?: number | null;
};

type SeatConstraint = {
  MinSeats?: number | null;
  MaxSeats?: number | null;
  Type?: string | null;
};

type LookbackWindow = {
  Type?: string;
  Period?: string;
};

type PrerequisiteRule = {
  Products?: ConstraintProduct[];
  SeatConstraint?: SeatConstraint;
  LookbackWindow?: LookbackWindow;
};

type ConstraintsData = {
  seatConstraints?: unknown[];
  assetOwnershipLimits?: unknown[];
  eligibilityConstraints?: unknown[];
  productOwnershipConstraints?: unknown[];
  purchaseRequirementConstraints?: unknown[];
  prerequisiteConstraints?: {
    MustHaveAll?: PrerequisiteRule | null;
    MustHaveAny?: PrerequisiteRule | null;
    MustHaveNone?: PrerequisiteRule | null;
  };
};

type EligibilityPolicy = {
  policyId?: string;
  policyData?: {
    eligibility?: {
      constraintsData?: ConstraintsData;
    };
  };
};

type GrowthMargin = {
  id: string;
  name: string;
  description: string;
  productCodes?: GrowthMarginProductCode[];
  requiredProducts?: RequiredProduct[];
  productPolicy?: {
    pricingPolicies?: PricingPolicy[];
    offerPolicies?: OfferPolicy[];
    eligibilityPolicies?: EligibilityPolicy[];
  };
};

type Customer = {
  id?: string;
  tenantId?: string;
  customerId?: string;
  companyProfile?: {
    companyName?: string;
    domain?: string;
    tenantId?: string;
  };
};

type EligibilityError = {
  type?: string;
  description?: string;
};

type IneligibilityGuidance = {
  reason: string;
  whatItMeans: string;
  whatToDo: string;
};

type SeatRequirement = {
  minSeats?: number;
  maxSeats?: number;
  type?: string;
  source: string;
};

type EligibilityResult = {
  priceBenefit?: {
    productCode?: string;
    productId?: string;
    skuId?: string;
    availabilityId?: string;
  };
  isEligible: boolean;
  errors?: EligibilityError[];
};

type EligibilityResponse = CollectionResponse<{
  id: number;
  eligibilities?: EligibilityResult[];
}>;

type EstimateItem = {
  id: string;
  growthMargin: GrowthMargin;
  requiredProductIndex: number;
  productCodeIndex: number;
  customerId: string;
  quantity: string;
  status: "idle" | "checking" | "eligible" | "ineligible" | "error";
  result?: EligibilityResult;
  error?: string;
};

type GrowthMarginCalculatorProps = {
  account: AccountInfo | null;
  requestJson: PartnerCenterApiClient;
};

function getCustomerId(customer: Customer) {
  return customer.id ?? customer.customerId ?? customer.tenantId ?? customer.companyProfile?.tenantId ?? "";
}

function getCustomerLabel(customer: Customer) {
  const id = getCustomerId(customer);
  const companyName = customer.companyProfile?.companyName;
  const domain = customer.companyProfile?.domain;

  return [companyName, domain, id].filter(Boolean).join(" - ") || "Unnamed customer";
}

function formatPercent(value?: number) {
  if (typeof value !== "number") {
    return "";
  }

  return `${Math.round(value * 10000) / 100}%`;
}

function getDiscountSummary(growthMargin: GrowthMargin) {
  const benefits = growthMargin.productPolicy?.pricingPolicies?.flatMap(
    (policy) => policy.policyData?.benefits ?? [],
  ) ?? [];

  if (benefits.length === 0) {
    return "Discount details unavailable";
  }

  return benefits
    .map((benefit) => {
      const value = benefit.type === "PercentDiscount" ? formatPercent(benefit.value) : benefit.value?.toString();
      return [value, benefit.type].filter(Boolean).join(" ");
    })
    .join(", ");
}

function getOfferWindow(growthMargin: GrowthMargin) {
  const properties = growthMargin.productPolicy?.offerPolicies?.[0]?.policyData?.properties;
  if (!properties?.startDate && !properties?.endDate) {
    return "Offer dates unavailable";
  }

  return `${properties.startDate ? new Date(properties.startDate).toLocaleDateString() : "Now"} - ${
    properties.endDate ? new Date(properties.endDate).toLocaleDateString() : "No end date"
  }`;
}

function parseBenefitId(id: string) {
  const [productId, skuId, availabilityId] = id.split(":");
  return { productId, skuId, availabilityId };
}

function buildEligibilityRequest(item: EstimateItem) {
  const requiredProduct = item.growthMargin.requiredProducts?.[item.requiredProductIndex];
  const productCode = item.growthMargin.productCodes?.[item.productCodeIndex]?.productCode;
  const benefitIdentity = parseBenefitId(item.growthMargin.id);

  if (!requiredProduct) {
    throw new Error("Select a required base product before checking eligibility.");
  }

  if (!productCode && (!benefitIdentity.productId || !benefitIdentity.skuId)) {
    throw new Error("This growth margin does not include enough identity information to check eligibility.");
  }

  const quantity = Number(item.quantity);
  if (!Number.isInteger(quantity) || quantity <= 0) {
    throw new Error("Enter a whole-number seat quantity greater than zero.");
  }

  return {
    items: [
      {
        id: 1,
        type: "GrowthMargin",
        evaluationType: "Immediate",
        targetProduct: {
          productId: requiredProduct.productId,
          skuId: requiredProduct.skuId,
          quantity,
          termDuration: requiredProduct.termDuration,
          billingCycle: requiredProduct.billingCycle,
        },
        priceBenefit: {
          ...(productCode ? { productCode } : {}),
          ...(benefitIdentity.productId ? { productId: benefitIdentity.productId } : {}),
          ...(benefitIdentity.skuId ? { skuId: benefitIdentity.skuId } : {}),
          ...(benefitIdentity.availabilityId ? { availabilityId: benefitIdentity.availabilityId } : {}),
        },
      },
    ],
  };
}

function getBestEligibilityResult(response: EligibilityResponse, item: EstimateItem) {
  const eligibilities = response.items?.[0]?.eligibilities ?? [];
  const productCode = item.growthMargin.productCodes?.[item.productCodeIndex]?.productCode;

  return (
    eligibilities.find((eligibility) => eligibility.priceBenefit?.productCode === productCode) ??
    eligibilities[0]
  );
}

function formatDateTime(value?: string) {
  return value ? new Date(value).toLocaleString() : "Not specified";
}

function formatSeatLimit(value: number | null | undefined) {
  return value === null || value === undefined ? "No limit" : value.toString();
}

function getPricingBenefits(growthMargin: GrowthMargin) {
  return growthMargin.productPolicy?.pricingPolicies?.flatMap((policy) =>
    (policy.policyData?.benefits ?? []).map((benefit) => ({
      policyId: policy.policyId ?? "Pricing policy",
      type: benefit.type ?? "Benefit",
      value: benefit.type === "PercentDiscount" ? formatPercent(benefit.value) : benefit.value?.toString() ?? "Not specified",
    })),
  ) ?? [];
}

function getPrerequisiteSummaries(growthMargin: GrowthMargin) {
  const policies = growthMargin.productPolicy?.eligibilityPolicies ?? [];
  return policies.flatMap((policy) => {
    const prerequisites = policy.policyData?.eligibility?.constraintsData?.prerequisiteConstraints;
    if (!prerequisites) {
      return [];
    }

    return ([
      ["Must have all", prerequisites.MustHaveAll],
      ["Must have any", prerequisites.MustHaveAny],
      ["Must not have", prerequisites.MustHaveNone],
    ] as const)
      .flatMap(([title, rule]) => {
        if (!rule) {
          return [];
        }

        return [{
          title,
          policyId: policy.policyId ?? "Eligibility policy",
          products: rule.Products ?? [],
          seatType: rule.SeatConstraint?.Type ?? "Not specified",
          minSeats: formatSeatLimit(rule.SeatConstraint?.MinSeats),
          maxSeats: formatSeatLimit(rule.SeatConstraint?.MaxSeats),
          lookback: rule.LookbackWindow?.Period ?? rule.LookbackWindow?.Type ?? "Not specified",
        }];
      });
  });
}

function getConstraintCounts(growthMargin: GrowthMargin) {
  const policies = growthMargin.productPolicy?.eligibilityPolicies ?? [];
  return policies.reduce(
    (counts, policy) => {
      const data = policy.policyData?.eligibility?.constraintsData;
      return {
        seat: counts.seat + (data?.seatConstraints?.length ?? 0),
        asset: counts.asset + (data?.assetOwnershipLimits?.length ?? 0),
        eligibility: counts.eligibility + (data?.eligibilityConstraints?.length ?? 0),
        productOwnership: counts.productOwnership + (data?.productOwnershipConstraints?.length ?? 0),
        purchase: counts.purchase + (data?.purchaseRequirementConstraints?.length ?? 0),
      };
    },
    { seat: 0, asset: 0, eligibility: 0, productOwnership: 0, purchase: 0 },
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function getNumericField(record: Record<string, unknown>, names: string[]) {
  for (const name of names) {
    const value = record[name];
    if (typeof value === "number") {
      return value;
    }
  }

  return undefined;
}

function getStringField(record: Record<string, unknown>, names: string[]) {
  for (const name of names) {
    const value = record[name];
    if (typeof value === "string") {
      return value;
    }
  }

  return undefined;
}

function collectSeatRequirements(value: unknown, source = "Growth margin constraint"): SeatRequirement[] {
  if (Array.isArray(value)) {
    return value.flatMap((entry) => collectSeatRequirements(entry, source));
  }

  if (!isRecord(value)) {
    return [];
  }

  const minSeats = getNumericField(value, ["MinSeats", "minSeats", "MinimumSeats", "minimumSeats"]);
  const maxSeats = getNumericField(value, ["MaxSeats", "maxSeats", "MaximumSeats", "maximumSeats"]);
  const type = getStringField(value, ["Type", "type", "constraintType", "seatConstraintType"]);
  const ownRequirement =
    minSeats !== undefined || maxSeats !== undefined
      ? [{ minSeats, maxSeats, type, source }]
      : [];

  return [
    ...ownRequirement,
    ...Object.entries(value).flatMap(([key, child]) => collectSeatRequirements(child, key)),
  ];
}

function getQuantityGuidance(item: EstimateItem) {
  const quantity = Number(item.quantity);
  if (!Number.isFinite(quantity)) {
    return null;
  }

  const constraintsData = item.growthMargin.productPolicy?.eligibilityPolicies?.flatMap((policy) => {
    const data = policy.policyData?.eligibility?.constraintsData;
    return data ? [data] : [];
  }) ?? [];
  const requirements = collectSeatRequirements(constraintsData)
    .filter((requirement) => requirement.minSeats !== undefined || requirement.maxSeats !== undefined)
    .filter((requirement) => {
      const source = requirement.source.toLowerCase();
      const type = (requirement.type ?? "").toLowerCase();
      return !source.includes("lookback") && !type.includes("basecumulativequantity");
    });
  const minimum = requirements
    .map((requirement) => requirement.minSeats)
    .filter((value): value is number => value !== undefined && value > 0)
    .sort((left, right) => left - right)[0];
  const maximum = requirements
    .map((requirement) => requirement.maxSeats)
    .filter((value): value is number => value !== undefined && value > 0)
    .sort((left, right) => right - left)[0];

  if (minimum !== undefined && quantity < minimum) {
    const delta = minimum - quantity;
    return `Current quantity is ${quantity}. This growth margin requires at least ${minimum} seats, so add ${delta} more seat${delta === 1 ? "" : "s"} and check eligibility again.`;
  }

  if (maximum !== undefined && quantity > maximum) {
    const delta = quantity - maximum;
    return `Current quantity is ${quantity}. This growth margin allows at most ${maximum} seats for this rule, so reduce the quantity by ${delta} seat${delta === 1 ? "" : "s"} or use a different growth margin.`;
  }

  return null;
}

function getIneligibilityGuidance(error: EligibilityError, item: EstimateItem): IneligibilityGuidance {
  const type = (error.type ?? "").toLowerCase();
  const description = (error.description ?? "").toLowerCase();
  const requiredProduct = item.growthMargin.requiredProducts?.[item.requiredProductIndex];

  if (type.includes("seat") || description.includes("minimum") || description.includes("maximum") || description.includes("seat")) {
    const quantityGuidance = getQuantityGuidance(item);
    return {
      reason: "Seat requirement not met",
      whatItMeans:
        "The entered quantity does not satisfy the growth margin seat rule. This usually maps to the documented Below minimum seats or Expansion multiple not met reasons.",
      whatToDo: quantityGuidance ?? `Adjust the quantity to meet the growth margin seat constraint, then check eligibility again${
        requiredProduct ? ` for ${requiredProduct.productId}:${requiredProduct.skuId}` : ""
      }. If this is a seat expansion scenario, use the required expansion multiple rather than a small incremental quantity.`,
    };
  }

  if (description.includes("new-to-offer") || description.includes("already has") || description.includes("already owns")) {
    return {
      reason: "Customer is not new to the offer",
      whatItMeans:
        "The customer appears to already have the product or a configured related SKU within the lookback period.",
      whatToDo:
        "Choose a customer that is new to this offer, or select a different growth margin/base SKU combination. The customer can still transact at standard pricing.",
    };
  }

  if (description.includes("existing subscription") || description.includes("seat quantity update")) {
    return {
      reason: "Existing subscription change is not eligible",
      whatItMeans:
        "Growth margins generally require a new subscription. Seat quantity updates on existing subscriptions are not evaluated for growth margin eligibility.",
      whatToDo:
        "Use a new subscription purchase flow, or only evaluate an existing subscription when it is a supported mid-term upgrade scenario.",
    };
  }

  if (description.includes("strategic sku") || description.includes("sku mix")) {
    return {
      reason: "Strategic SKU mix requirement not met",
      whatItMeans:
        "The tenant does not meet the configured strategic SKU mix threshold, or it is already above the threshold.",
      whatToDo:
        "Review the customer's current SKU mix and choose a qualifying base SKU or customer scenario before checking again.",
    };
  }

  if (description.includes("channel shift") || description.includes("ea to csp")) {
    return {
      reason: "Channel-shift seats are excluded",
      whatItMeans:
        "Seats that originate from a channel shift, such as EA to CSP, do not qualify as growth for this margin.",
      whatToDo:
        "Use net-new eligible seats/customers rather than channel-shifted seats, or proceed without the growth margin.",
    };
  }

  if (description.includes("specialized offer")) {
    return {
      reason: "Specialized Offer takes precedence",
      whatItMeans:
        "A Specialized Offer applies to the SKU and takes precedence over the growth margin.",
      whatToDo:
        "Review the available price benefits for the SKU. The growth margin may not apply while the Specialized Offer is active.",
    };
  }

  return {
    reason: error.type ?? "Eligibility rule not met",
    whatItMeans: error.description ?? "Partner Center returned this item as not eligible for the selected growth margin.",
    whatToDo:
      "Review the selected customer, base product, quantity, term, and billing cycle. If the inputs are correct, proceed at standard pricing because growth margin ineligibility is non-blocking.",
  };
}

function GrowthMarginDetailDialog({
  growthMargin,
  onClose,
}: {
  growthMargin: GrowthMargin;
  onClose: () => void;
}) {
  const pricingBenefits = getPricingBenefits(growthMargin);
  const prerequisiteSummaries = getPrerequisiteSummaries(growthMargin);
  const constraintCounts = getConstraintCounts(growthMargin);
  const offerPolicies = growthMargin.productPolicy?.offerPolicies ?? [];
  const productCodes = growthMargin.productCodes ?? [];
  const requiredProducts = growthMargin.requiredProducts ?? [];

  return (
    <section className="growth-margin-detail-modal" role="dialog" aria-modal="true" aria-label="Growth margin details">
      <div className="modal-header">
        <div>
          <p className="eyebrow">Growth margin details</p>
          <h2>{growthMargin.name}</h2>
        </div>
        <button className="ghost" onClick={onClose}>
          Close
        </button>
      </div>

      <div className="detail-overview">
        <div className="detail-hero-card">
          <span>Benefit</span>
          <strong>{getDiscountSummary(growthMargin)}</strong>
          <p>{growthMargin.description}</p>
        </div>
        <div className="detail-hero-card">
          <span>Offer window</span>
          <strong>{getOfferWindow(growthMargin)}</strong>
          <p>Growth margins are partner-earned economics and are validated at transaction time.</p>
        </div>
      </div>

      <div className="detail-section-grid">
        <section className="detail-section">
          <div className="detail-section-heading">
            <h3>Growth margin identity</h3>
            <a
              href="https://learn.microsoft.com/en-us/partner-center/developer/get-growth-margins"
              target="_blank"
              rel="noreferrer"
            >
              API docs
            </a>
          </div>
          <dl className="detail-definition-list">
            <div>
              <dt>Catalog ID</dt>
              <dd>{growthMargin.id}</dd>
            </div>
            <div>
              <dt>Product codes</dt>
              <dd>
                {productCodes.length
                  ? productCodes.map((code) => `${code.productCode}${code.termId ? ` (${code.termId})` : ""}`).join(", ")
                  : "Not returned"}
              </dd>
            </div>
          </dl>
        </section>

        <section className="detail-section">
          <h3>Pricing policy</h3>
          {pricingBenefits.length === 0 ? (
            <p className="empty">No pricing benefits were returned.</p>
          ) : (
            <div className="detail-table">
              <div className="detail-table-row head">
                <span>Policy</span>
                <span>Type</span>
                <span>Value</span>
              </div>
              {pricingBenefits.map((benefit, index) => (
                <div className="detail-table-row" key={`${benefit.policyId}-${index}`}>
                  <span>{benefit.policyId}</span>
                  <span>{benefit.type}</span>
                  <strong>{benefit.value}</strong>
                </div>
              ))}
            </div>
          )}
        </section>

        <section className="detail-section wide">
          <h3>Required base products</h3>
          {requiredProducts.length === 0 ? (
            <p className="empty">No required products were returned.</p>
          ) : (
            <div className="detail-table required-products-table">
              <div className="detail-table-row head">
                <span>Product ID</span>
                <span>SKU ID</span>
                <span>Term</span>
                <span>Billing cycle</span>
              </div>
              {requiredProducts.map((product, index) => (
                <div className="detail-table-row" key={`${product.productId}-${product.skuId}-${index}`}>
                  <span>{product.productId}</span>
                  <span>{product.skuId}</span>
                  <span>{product.termDuration}</span>
                  <span>{product.billingCycle}</span>
                </div>
              ))}
            </div>
          )}
        </section>

        <section className="detail-section">
          <h3>Offer policy</h3>
          {offerPolicies.length === 0 ? (
            <p className="empty">No offer policies were returned.</p>
          ) : (
            <div className="policy-card-list">
              {offerPolicies.map((policy, index) => (
                <article className="policy-card" key={`${policy.policyId ?? "offer"}-${index}`}>
                  <strong>{policy.policyId ?? "Offer policy"}</strong>
                  <span>Beneficiary: {policy.policyData?.beneficiary ?? "Not specified"}</span>
                  <span>Start: {formatDateTime(policy.policyData?.properties?.startDate)}</span>
                  <span>End: {formatDateTime(policy.policyData?.properties?.endDate)}</span>
                  <span>{policy.policyData?.isOptional ? "Optional" : "Automatically evaluated"}</span>
                </article>
              ))}
            </div>
          )}
        </section>

        <section className="detail-section">
          <h3>Eligibility constraints</h3>
          <div className="constraint-count-grid">
            <span>Seat constraints <strong>{constraintCounts.seat}</strong></span>
            <span>Asset ownership <strong>{constraintCounts.asset}</strong></span>
            <span>Eligibility rules <strong>{constraintCounts.eligibility}</strong></span>
            <span>Product ownership <strong>{constraintCounts.productOwnership}</strong></span>
            <span>Purchase requirements <strong>{constraintCounts.purchase}</strong></span>
            <span>Prerequisites <strong>{prerequisiteSummaries.length}</strong></span>
          </div>

          {prerequisiteSummaries.length > 0 && (
            <div className="prerequisite-list">
              {prerequisiteSummaries.map((rule, index) => (
                <article className="policy-card" key={`${rule.title}-${index}`}>
                  <strong>{rule.title}</strong>
                  <span>{rule.policyId}</span>
                  <span>
                    Products:{" "}
                    {rule.products.length
                      ? rule.products
                          .map(
                            (product) =>
                              `${product.BigId ?? "Unknown product"} (min ${formatSeatLimit(product.MinSeats)}, max ${formatSeatLimit(product.MaxSeats)})`,
                          )
                          .join(", ")
                      : "None listed"}
                  </span>
                  <span>Seat rule: {rule.seatType} (min {rule.minSeats}, max {rule.maxSeats})</span>
                  <span>Lookback: {rule.lookback}</span>
                </article>
              ))}
            </div>
          )}
        </section>
      </div>
    </section>
  );
}

export function GrowthMarginCalculator({ account, requestJson }: GrowthMarginCalculatorProps) {
  const [country, setCountry] = useState("US");
  const [segment, setSegment] = useState("commercial");
  const [baseSkuPaths, setBaseSkuPaths] = useState("");
  const [search, setSearch] = useState("");
  const [activeFilter, setActiveFilter] = useState("Featured");
  const [growthMargins, setGrowthMargins] = useState<GrowthMargin[]>([]);
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [estimateItems, setEstimateItems] = useState<EstimateItem[]>([]);
  const [detailGrowthMargin, setDetailGrowthMargin] = useState<GrowthMargin | null>(null);
  const [growthMarginsLoading, setGrowthMarginsLoading] = useState(false);
  const [customersLoading, setCustomersLoading] = useState(false);
  const [message, setMessage] = useState("");
  const autoLoadedAccountRef = useRef("");

  const customerOptions = useMemo(
    () => customers.filter((customer) => getCustomerId(customer)),
    [customers],
  );
  const visibleGrowthMargins = useMemo(() => {
    const query = search.trim().toLowerCase();
    return growthMargins.filter((growthMargin) => {
      const matchesSearch = !query || [
        growthMargin.name,
        growthMargin.description,
        growthMargin.id,
        growthMargin.productCodes?.map((code) => code.productCode).join(" "),
        growthMargin.requiredProducts?.map((product) => `${product.productId} ${product.skuId}`).join(" "),
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(query);

      if (!matchesSearch) {
        return false;
      }

      if (activeFilter === "Annual") {
        return growthMargin.requiredProducts?.some((product) => product.billingCycle.toLowerCase() === "annual");
      }

      if (activeFilter === "Monthly") {
        return growthMargin.requiredProducts?.some((product) => product.billingCycle.toLowerCase() === "monthly");
      }

      if (activeFilter === "Multiple products") {
        return (growthMargin.requiredProducts?.length ?? 0) > 1;
      }

      if (activeFilter === "High margin") {
        return growthMargin.productPolicy?.pricingPolicies?.some((policy) =>
          policy.policyData?.benefits?.some((benefit) => (benefit.value ?? 0) >= 0.1),
        );
      }

      return true;
    });
  }, [activeFilter, growthMargins, search]);
  const filterCounts = useMemo(() => {
    const annual = growthMargins.filter((growthMargin) =>
      growthMargin.requiredProducts?.some((product) => product.billingCycle.toLowerCase() === "annual"),
    ).length;
    const monthly = growthMargins.filter((growthMargin) =>
      growthMargin.requiredProducts?.some((product) => product.billingCycle.toLowerCase() === "monthly"),
    ).length;
    const multipleProducts = growthMargins.filter((growthMargin) => (growthMargin.requiredProducts?.length ?? 0) > 1).length;
    const highMargin = growthMargins.filter((growthMargin) =>
      growthMargin.productPolicy?.pricingPolicies?.some((policy) =>
        policy.policyData?.benefits?.some((benefit) => (benefit.value ?? 0) >= 0.1),
      ),
    ).length;

    return {
      Featured: growthMargins.length,
      Annual: annual,
      Monthly: monthly,
      "Multiple products": multipleProducts,
      "High margin": highMargin,
    };
  }, [growthMargins]);
  const eligibleCount = estimateItems.filter((item) => item.status === "eligible").length;
  const ineligibleCount = estimateItems.filter((item) => item.status === "ineligible").length;

  useEffect(() => {
    const accountKey = account?.homeAccountId ?? account?.username ?? "";
    if (!accountKey || autoLoadedAccountRef.current === accountKey) {
      return;
    }

    autoLoadedAccountRef.current = accountKey;
    void loadGrowthMargins();
    void loadCustomers();
  }, [account]);

  async function loadGrowthMargins() {
    setGrowthMarginsLoading(true);
    setMessage("");
    trackClarityEvent("growth_margins_load_clicked");

    try {
      const params = new URLSearchParams({
        type: "growthmargin",
        country: country.trim().toUpperCase(),
        segment: segment.trim(),
      });

      if (baseSkuPaths.trim()) {
        params.set("baseSkuPaths", baseSkuPaths.trim());
      }

      const response = await requestJson<CollectionResponse<GrowthMargin>>(`/catalog/benefits?${params}`);
      setGrowthMargins(response.items ?? []);
      trackClarityEvent("growth_margins_loaded");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Could not load growth margins.");
      trackClarityEvent("growth_margins_load_failed");
    } finally {
      setGrowthMarginsLoading(false);
    }
  }

  async function loadCustomers() {
    setCustomersLoading(true);
    setMessage("");
    trackClarityEvent("customers_load_clicked");

    try {
      const response = await requestJson<CollectionResponse<Customer>>("/customers");
      setCustomers(response.items ?? []);
      trackClarityEvent("customers_loaded");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Could not load customers.");
      trackClarityEvent("customers_load_failed");
    } finally {
      setCustomersLoading(false);
    }
  }

  function addToEstimate(growthMargin: GrowthMargin) {
    trackClarityEvent("growth_margin_added_to_estimate");
    setEstimateItems((current) => [
      ...current,
      {
        id: crypto.randomUUID(),
        growthMargin,
        requiredProductIndex: 0,
        productCodeIndex: 0,
        customerId: customerOptions[0] ? getCustomerId(customerOptions[0]) : "",
        quantity: growthMargin.requiredProducts?.[0] ? "1" : "",
        status: "idle",
      },
    ]);
  }

  function updateEstimateItem(id: string, updates: Partial<EstimateItem>) {
    setEstimateItems((current) =>
      current.map((item) =>
        item.id === id
          ? {
              ...item,
              ...updates,
              status: updates.status ?? "idle",
              result: updates.result,
              error: updates.error,
            }
          : item,
      ),
    );
  }

  function removeEstimateItem(id: string) {
    setEstimateItems((current) => current.filter((item) => item.id !== id));
  }

  async function checkEligibility(item: EstimateItem) {
    if (!item.customerId) {
      updateEstimateItem(item.id, { status: "error", error: "Select a customer before checking eligibility." });
      return;
    }

    updateEstimateItem(item.id, { status: "checking" });
    trackClarityEvent("eligibility_check_started");

    try {
      const requestBody = buildEligibilityRequest(item);
      const response = await requestJson<EligibilityResponse>(
        `/customers/${encodeURIComponent(item.customerId)}/priceBenefitEligibilities`,
        {
          method: "POST",
          body: JSON.stringify(requestBody),
        },
      );
      const result = getBestEligibilityResult(response, item);

      if (!result) {
        updateEstimateItem(item.id, {
          status: "ineligible",
          error: "No eligibility result was returned for this growth margin.",
        });
        return;
      }

      updateEstimateItem(item.id, {
        status: result.isEligible ? "eligible" : "ineligible",
        result,
      });
      trackClarityEvent("eligibility_check_completed");
      trackClarityEvent(result.isEligible ? "eligibility_result_eligible" : "eligibility_result_ineligible");
    } catch (error) {
      updateEstimateItem(item.id, {
        status: "error",
        error: error instanceof Error ? error.message : "Eligibility check failed.",
      });
      trackClarityEvent("eligibility_check_failed");
    }
  }

  return (
    <section className="calculator-stack">
      <section className="pricing-summary-bar">
        <div className="estimate-title-group">
          <h1>Your Estimate</h1>
          <span>Growth Margin Calculator</span>
        </div>
        <div className="estimate-metrics" aria-label="Estimate summary">
          <div>
            <strong>{estimateItems.length}</strong>
            <span>Growth margin items</span>
          </div>
          <div>
            <strong>{eligibleCount}</strong>
            <span>Eligible checks</span>
          </div>
          <div>
            <strong>{ineligibleCount}</strong>
            <span>Not eligible</span>
          </div>
        </div>
        <button className="ghost" onClick={loadCustomers} disabled={!account || customersLoading}>
          {customersLoading ? "Loading..." : customerOptions.length ? "Refresh customers" : "Load customers"}
        </button>
      </section>

      {message && <p className="message danger">{message}</p>}

      <section className="pricing-calculator-shell">
        <section className="pricing-products-area">
          <div className="pricing-tabs" role="tablist" aria-label="Growth margin calculator sections">
            <button className="pricing-tab active">Growth Margin Products</button>
          </div>

          <div className="pricing-instruction-bar">
            Select a growth margin to include it in your estimate.
          </div>

          <div className="pricing-toolbar">
            <input
              className="input pricing-search clarity-mask"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search growth margins, product ID, SKU, or product code"
            />
            <div className="pricing-toolbar-fields">
              <label>
                Country
                <input
                  className="input clarity-mask"
                  value={country}
                  onChange={(event) => setCountry(event.target.value)}
                  maxLength={2}
                />
              </label>
              <label>
                Segment
                <select className="input" value={segment} onChange={(event) => setSegment(event.target.value)}>
                  <option value="commercial">commercial</option>
                </select>
              </label>
              <label>
                Optional base SKU paths
                <input
                  className="input clarity-mask"
                  value={baseSkuPaths}
                  onChange={(event) => setBaseSkuPaths(event.target.value)}
                  placeholder="CFQ7TTC0ZSXK:0002,CFQ7TTC0ZSXK:0005"
                />
              </label>
              <button className="primary" onClick={loadGrowthMargins} disabled={!account || growthMarginsLoading}>
                {growthMarginsLoading ? "Loading..." : "Load growth margins"}
              </button>
            </div>
          </div>

          <aside className="pricing-category-list" aria-label="Growth margin filters">
            {Object.entries(filterCounts).map(([filter, count]) => (
              <button
                className={activeFilter === filter ? "active" : ""}
                key={filter}
                onClick={() => setActiveFilter(filter)}
              >
                <span>{filter}</span>
                <strong>{count}</strong>
              </button>
            ))}
          </aside>

          <div className="pricing-products-grid">
            <div className="growth-margin-product-grid clarity-mask">
              {growthMargins.length === 0 && (
                <p className="empty">
                  Load growth margins to see active offers for the selected market and segment.
                </p>
              )}
              {growthMargins.length > 0 && visibleGrowthMargins.length === 0 && (
                <p className="empty">No growth margins match the current filter.</p>
              )}
              {visibleGrowthMargins.map((growthMargin) => (
                <article className="pricing-product-card" key={growthMargin.id}>
                  <div className="pricing-product-icon">GM</div>
                  <div className="pricing-product-copy">
                    <h3>{growthMargin.name}</h3>
                    <p>{growthMargin.description}</p>
                  </div>
                  <div className="pricing-product-meta">
                    <span>{getDiscountSummary(growthMargin)}</span>
                    <span>{growthMargin.requiredProducts?.length ?? 0} required product(s)</span>
                  </div>
                  <div className="pricing-card-actions">
                    <button
                      className="ghost"
                      onClick={() => {
                        setDetailGrowthMargin(growthMargin);
                        trackClarityEvent("growth_margin_details_opened");
                      }}
                    >
                      Details
                    </button>
                    <button className="primary" onClick={() => addToEstimate(growthMargin)}>
                      Add to estimate
                    </button>
                  </div>
                </article>
              ))}
            </div>
          </div>
        </section>

        <section className="estimate-panel pricing-estimate-drawer clarity-mask">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">Estimate</p>
              <h2>Your estimate</h2>
            </div>
            <span className="pill">{estimateItems.length}</span>
          </div>

          <div className="estimate-list">
            {estimateItems.length === 0 && (
              <p className="empty">
                Add a growth margin to build an estimate and check customer eligibility.
              </p>
            )}

            {estimateItems.map((item, index) => {
              const requiredProducts = item.growthMargin.requiredProducts ?? [];
              const productCodes = item.growthMargin.productCodes ?? [];
              const selectedResult = item.result;
              const errors = selectedResult?.errors ?? [];

              return (
                <article className="estimate-card" key={item.id}>
                  <div className="estimate-card-header">
                    <div>
                      <span className="pill">Item {index + 1}</span>
                      <h3>{item.growthMargin.name}</h3>
                    </div>
                    <button className="ghost" onClick={() => removeEstimateItem(item.id)}>
                      Remove
                    </button>
                  </div>

                  <div className="estimate-form-grid">
                    <label>
                      Customer
                      <select
                        className="input"
                        value={item.customerId}
                        onChange={(event) => updateEstimateItem(item.id, { customerId: event.target.value })}
                      >
                        <option value="">Select customer</option>
                        {customerOptions.map((customer) => {
                          const customerId = getCustomerId(customer);
                          return (
                            <option value={customerId} key={customerId}>
                              {getCustomerLabel(customer)}
                            </option>
                          );
                        })}
                      </select>
                    </label>

                    <label>
                      Quantity
                      <input
                        className="input"
                        type="number"
                        min="1"
                        value={item.quantity}
                        onChange={(event) => updateEstimateItem(item.id, { quantity: event.target.value })}
                      />
                    </label>

                    <label>
                      Required base product
                      <select
                        className="input"
                        value={item.requiredProductIndex}
                        onChange={(event) =>
                          updateEstimateItem(item.id, { requiredProductIndex: Number(event.target.value) })
                        }
                      >
                        {requiredProducts.length === 0 && <option value={0}>No required product returned</option>}
                        {requiredProducts.map((product, productIndex) => (
                          <option value={productIndex} key={`${product.productId}-${product.skuId}-${productIndex}`}>
                            {product.productId}:{product.skuId} - {product.termDuration} - {product.billingCycle}
                          </option>
                        ))}
                      </select>
                    </label>

                    <label>
                      Growth margin product code
                      <select
                        className="input"
                        value={item.productCodeIndex}
                        onChange={(event) =>
                          updateEstimateItem(item.id, { productCodeIndex: Number(event.target.value) })
                        }
                      >
                        {productCodes.length === 0 && <option value={0}>Use growth margin ID</option>}
                        {productCodes.map((productCode, productCodeIndex) => (
                          <option value={productCodeIndex} key={`${productCode.productCode}-${productCodeIndex}`}>
                            {productCode.productCode}
                            {productCode.termId ? ` - ${productCode.termId}` : ""}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>

                  <div className="estimate-actions">
                    <button
                      className="primary"
                      onClick={() => checkEligibility(item)}
                      disabled={!account || item.status === "checking"}
                    >
                      {item.status === "checking" ? "Checking..." : "Check eligibility"}
                    </button>
                    <span className={`eligibility-badge ${item.status}`}>
                      {item.status === "eligible"
                        ? "Eligible"
                        : item.status === "ineligible"
                          ? "Not eligible"
                          : item.status === "error"
                            ? "Needs attention"
                            : item.status === "checking"
                              ? "Checking"
                              : "Not checked"}
                    </span>
                  </div>

                  {(selectedResult || item.error) && (
                    <div className="eligibility-result">
                      {item.error && <p>{item.error}</p>}
                      {selectedResult?.isEligible && (
                        <p>
                          This customer is eligible for the selected growth margin. Confirm the benefit on the cart
                          line item before submitting a transaction.
                        </p>
                      )}
                      {!selectedResult?.isEligible && errors.length > 0 && (
                        <div className="ineligibility-guidance-list">
                          <strong>Why this is not eligible</strong>
                          {errors.map((error, errorIndex) => {
                            const guidance = getIneligibilityGuidance(error, item);
                            return (
                              <article className="ineligibility-guidance-card" key={`${error.type ?? "error"}-${errorIndex}`}>
                                <h4>{guidance.reason}</h4>
                                <p>{guidance.whatItMeans}</p>
                                <div>
                                  <span>What to do</span>
                                  <p>{guidance.whatToDo}</p>
                                </div>
                                {error.description && <small>Partner Center message: {error.description}</small>}
                              </article>
                            );
                          })}
                        </div>
                      )}
                      {!selectedResult?.isEligible && errors.length === 0 && !item.error && (
                        <p>
                          The customer is not eligible for this growth margin. The transaction can still proceed at
                          standard pricing.
                        </p>
                      )}
                    </div>
                  )}
                </article>
              );
            })}
          </div>
        </section>
      </section>
      {detailGrowthMargin && (
        <div className="modal-backdrop" role="presentation">
          <div className="clarity-mask">
            <GrowthMarginDetailDialog growthMargin={detailGrowthMargin} onClose={() => setDetailGrowthMargin(null)} />
          </div>
        </div>
      )}
    </section>
  );
}
