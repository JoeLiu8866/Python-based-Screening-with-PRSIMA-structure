from scipy.stats import norm

def calculate_sample_size(N, confidence_level, p=0.5, E=0.05):
    # Z-value based on the desired confidence level
    Z = norm.ppf(1 - (1 - confidence_level) / 2)
    
    # 1 - p
    q = 1 - p
    
    # Calculate the sample size using Neyman's formula
    numerator = N * Z**2 * p * q
    denominator = E**2 * (N - 1) + Z**2 * p * q
    n = numerator / denominator
    
    return n

def main():
    # Input from the user
    N = int(input("Enter the total population size (N): "))
    confidence_level = float(input("Enter the confidence level (e.g., 0.95 for 95% confidence): "))
    p = float(input("Enter the population proportion (p, default is 0.5): ") or 0.5)
    E = float(input("Enter the margin of error (E): "))
    
    # Calculate the sample size
    sample_size = calculate_sample_size(N, confidence_level, p, E)
    
    # Output the result
    print(f"The required sample size is: {int(sample_size)}")

if __name__ == "__main__":
    main()
