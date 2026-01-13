# Pioneer Analytics - Grower Data Dashboard

Private internal tool for analyzing 5 years of grower data and forecasting 2027 sales projections.

## Features

- **5-Year Historical Analysis**: Visualize revenue, invoices, and growth trends from 2022-2026
- **2027 Forecasting**: Predictive models with confidence intervals (conservative, likely, optimistic)
- **Grower Analysis**: Top performers, retention rates, new vs returning customers
- **Product Mix**: Category breakdowns and product trend analysis
- **Data Management**: CSV upload, manual entry, and bulk import
- **Export Reports**: Download complete analytics reports

## Quick Start (Local Development)

1. Open `index.html` directly in your browser for client-side only mode (data stored in localStorage)
2. Default access code: `pioneer2024`

## Full Server Setup

```bash
cd internal-pioneer/server
npm install
cp .env.example .env
# Edit .env with your MongoDB URI and secrets
npm start
```

Access at `http://localhost:3001`

## Data Format

### CSV Upload
```
date,invoice_number,grower_name,product,quantity,amount
2024-03-15,INV-001,Smith Farm,Corn Seed,500,12500.00
2024-03-16,INV-002,Jones Acres,Soybean Seed,300,8750.00
```

### Supported Products
- Corn Seed
- Soybean Seed
- Herbicide
- Fungicide
- Insecticide
- Fertilizer
- Equipment
- Other

## Forecasting Methodology

The 2027 projections use a combination of:
1. **Linear Trend Analysis**: Projects based on historical slope
2. **Growth Rate Projection**: Applies average year-over-year growth
3. **Weighted Averaging**: Recent years weighted more heavily
4. **Seasonality Patterns**: Monthly distribution based on historical patterns

Confidence intervals are calculated based on historical volatility.

## Environment Variables

| Variable | Description |
|----------|-------------|
| MONGODB_URI | MongoDB connection string |
| JWT_SECRET | Secret for JWT token signing |
| ACCESS_CODE | Dashboard access password |
| PORT | Server port (default: 3001) |

## Security

- Password-protected access
- JWT authentication for API calls
- No data stored externally in client-only mode
- All communications over HTTPS in production

## Deployment

Deploy to Render using the included `render.yaml` configuration.
