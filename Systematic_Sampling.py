import numpy as np

def systematic_sampling(total_population, sample_size):
    # Calculate sampling interval
    interval = total_population / sample_size
    
    # Randomly select a starting point, starting from 1 instead of 0
    start = np.random.randint(1, int(interval) + 1)
    
    # Create a list of sample indices
    sample_indices = [int(start + i * interval) for i in range(sample_size)]
    
    # Ensure indices are within the total population range
    sample_indices = [(i - 1) % total_population + 1 for i in sample_indices]
    
    return sample_indices

def main():
    try:
        # Get user input
        total_population = int(input("Please enter the total population: "))
        sample_size = int(input("Please enter the sample size: "))
        
        # Ensure sample size does not exceed total population
        if sample_size > total_population:
            raise ValueError("Sample size cannot exceed total population")
        
        # Get sample indices
        sample_indices = systematic_sampling(total_population, sample_size)
        
        # Output sample indices
        print("The selected indices are: ", sample_indices)
        
    except ValueError as ve:
        print("Input error: ", ve)

if __name__ == "__main__":
    main()







