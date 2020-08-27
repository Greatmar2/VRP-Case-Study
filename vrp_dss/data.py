"""This script provides functionality to import data and convert it to usable forms."""
import json
from math import inf
from typing import List, Optional, Tuple, Union, Dict

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from settings import Data
from travel_matrix import read_input_data


class VehicleType:
    """Stores information about a vehicle type."""

    def __init__(self, data_index: int, name: str, distance_cost: float, time_cost: float, capacity: int,
                 vehicles_available: int):
        self.data_index = data_index
        self.name = name
        self.distance_cost = distance_cost
        self.time_cost = time_cost
        self.capacity = capacity
        self.vehicles_available = vehicles_available
        self.vehicles_used = 0

    def __eq__(self, other):
        if isinstance(other, VehicleType):
            return other.data_index == self.data_index or other.name == self.name
        if isinstance(other, Vehicle):
            return other.vehicle_type == self
        return False


class Vehicle:
    """Stores information about a vehicle."""

    def __init__(self, name: str, vehicle_type: VehicleType):
        self.name = name
        self.vehicle_type = vehicle_type

    def __eq__(self, other):
        if isinstance(other, Vehicle):
            return other.name == self.name
        return False


class PhysicalLocation:
    """Stores information about a real location serviced by SPAR. Each location can have a handful of customers,
    as SPAR, SUPERSPAR, KWIKSPAR, TOPS, and PHARMACY at the same location are treated as separate customers by SPAR.

    This script will treat stores at the same location as the same customer."""

    def __init__(self, data_index: int, name: str, latitude: float, longitude: float, offload: Union[str, float],
                 stores: str):
        """Initialise the object and convert the offload time and stores strings into appropriate values."""
        # super().__init__(data_index)
        self.data_index = data_index
        self.name = name
        self.latitude = latitude
        self.longitude = longitude
        if isinstance(offload, float):
            self.average_offload_time = offload
        else:
            self.average_offload_time = self.extract_average_offload(offload)
        self.store_ids = self.extract_store_ids(stores)
        self.demand = 0

    def __index__(self):
        return self.data_index

    def __repr__(self):
        return f"Location {self.name}, with store IDs {self.store_ids}."

    def __eq__(self, other):
        if isinstance(other, PhysicalLocation):
            return other.name == self.name
        if isinstance(other, str):
            return self.store_ids.count(other) > 0
        return False

    @property
    def anonymous_name(self) -> str:
        return f"{self.data_index}-{self.name[0]}"

    @staticmethod
    def extract_average_offload(offload_str: str) -> float:
        """
        Extracts the average offload time from the "target offload time" string.

        :param offload_str: String containing the target offload time and the target offload time.
        :returns: Integer representing the average offload time in hours.
        """
        # Expected offload string format: 00:00  (16:58)   min/container
        timestamp = offload_str[offload_str.index("(") + 1:offload_str.index(")")]
        time_units = timestamp.split(":")
        offload_time = (int(time_units[0]) * 60 + int(time_units[1])) / 3600

        return offload_time

    @staticmethod
    def extract_store_ids(stores_str: str) -> List[str]:
        """Extracts a list of the store IDs from the "stores at location" string.

        :param stores_str: String containing a list of stores at the location.
        :returns: A list of the IDs of stores at that location."""
        # Expected store string format: (3) 35668-Aurora SPAR, 35805-Aurora TOPS, 35924-Aurora PHARMACY
        stores = stores_str[4:].split(", ")
        store_ids: List[str] = []
        for store in stores:
            store_ids.append(store.split("-")[0])

        return store_ids


class Stop:
    def __init__(self, location: PhysicalLocation, delivered: int):
        self.location = location
        self.delivered = delivered


class Route:
    """Stores information from the archive about a route that a vehicle travelled."""

    def __init__(self, code: int, vehicle: Vehicle):
        self.code = code
        self.vehicle = vehicle
        self.stops: List[Stop] = []

    def __eq__(self, other):
        if isinstance(other, Route):
            return other.code == self.code
        if isinstance(other, int):
            return other == self.code
        return False

    def to_list(self) -> Tuple[int, List[Tuple[int, int]]]:
        """Returns a tuple containing list representing the route that can be used"""
        route: List[Tuple[int, int]] = [(stop.location.data_index, stop.delivered) for stop in self.stops]
        return self.vehicle.vehicle_type.data_index, route


def read_locations(filename: str) -> List[PhysicalLocation]:
    """Reads in the store locations and creates a list of PhysicalLocation objects."""
    workbook = load_workbook(filename=filename, read_only=True)
    locations_sheet = workbook["Store Locations"]
    locations: List[PhysicalLocation] = [PhysicalLocation(0, "Depot", -34.005993, 18.537626, 0.09167, "")]
    row = 2
    while locations_sheet[f"A{row}"].value:
        locations.append(PhysicalLocation(data_index=row - 2, name=locations_sheet[f"A{row}"].value,
                                          latitude=locations_sheet[f"B{row}"].value,
                                          longitude=locations_sheet[f"C{row}"].value,
                                          offload=locations_sheet[f"E{row}"].value,
                                          stores=locations_sheet[f"F{row}"].value))
        row += 1
    return locations


def find_location(store_id: str, locations: List[PhysicalLocation]) -> Optional[PhysicalLocation]:
    """Searches the list of locations for one that contains the given store ID."""
    # Hopefully List.index can utilise the custom __eq__ function for PhysicalLocation.
    # Otherwise I'll use a list comprehension.
    try:
        return locations[locations.index(store_id)]
    except ValueError:
        return None


def read_vehicles(types_file: str, types_sheet: str, vehicles_file: str, vehicles_sheet: str) -> Tuple[
    List[VehicleType], List[Vehicle]]:
    """Reads in information about the vehicle types and vehicles."""
    types_workbook = load_workbook(filename=types_file, read_only=True)
    types_sheet = types_workbook[types_sheet]
    vehicle_types: List[VehicleType] = []
    row = 2
    while types_sheet[f"A{row}"].value:
        vehicle_types.append(VehicleType(row - 2, name=types_sheet[f"A{row}"].value,
                                         distance_cost=types_sheet[f"B{row}"].value,
                                         time_cost=types_sheet[f"C{row}"].value,
                                         capacity=types_sheet[f"D{row}"].value,
                                         vehicles_available=types_sheet[f"E{row}"].value))
        row += 1

    vehicles_workbook = load_workbook(filename=vehicles_file, read_only=True)
    vehicles_sheet = vehicles_workbook[vehicles_sheet]
    vehicles: List[Vehicle] = []
    row = 2
    capacities = [vehicle_type.capacity for vehicle_type in vehicle_types]
    while vehicles_sheet[f"A{row}"].value:
        try:
            capacity = int(vehicles_sheet[f"C{row}"].value)
            vehicles.append(Vehicle(name=vehicles_sheet[f"A{row}"].value,
                                    vehicle_type=vehicle_types[capacities.index(capacity)]))
        except ValueError:
            pass
        row += 1

    return vehicle_types, vehicles


def find_vehicle(horse: str, trailer: str, carried: int, vehicles: List[Vehicle],
                 vehicle_types: List[VehicleType]) -> Optional[Vehicle]:
    """Searches the list of vehicles for one that contains the given vehicle name."""
    try:
        # Check if there is a vehicle that matches the horse code
        return vehicles[vehicles.index(horse)]
    except ValueError:
        try:
            # Check if there is a vehicle that matches the trailer code
            return vehicles[vehicles.index(trailer)]
        except ValueError:
            try:
                # Otherwise, try find the vehicle type that has the most similar to the capacity to the load carried
                lowest_diff = inf
                closest_type = None
                for vehicle_type in vehicle_types:
                    diff = abs(vehicle_type.capacity - carried)
                    if diff < lowest_diff:
                        lowest_diff = diff
                        closest_type = vehicle_type

                return Vehicle(horse, closest_type)

            except ValueError:
                return None


def find_route(route_code: str, routes: Dict[int, List[Route]]) -> Optional[Route]:
    """Searches the list of routes for one that contains the given route code."""
    for vehicle_type_index, tour in routes.items():
        count = tour.count(route_code)
        if count > 0:
            return tour[tour.index(route_code)]
        else:
            return None


def read_archive(filename: str, locations: List[PhysicalLocation], vehicles: List[Vehicle],
                 vehicle_types: List[VehicleType]) -> Dict[int, List[Route]]:
    """Reads in the delivery archive and sums up the demand per location and creating route lists."""
    workbook = load_workbook(filename=filename, read_only=True)
    deliveries_sheet = workbook["Deliveries"]
    routes: Dict[int, List[Route]] = {}
    row = 2
    route: Optional[Route] = None
    while deliveries_sheet[f"A{row}"].value:
        # Get the route code
        code = deliveries_sheet[f"D{row}"].value
        # If the route code does not match, then the vehicle is probably finished
        if code != route.code:
            if route:
                # In the archived data, vehicles are sometimes loaded with more than 30 pallets.
                # The schedulers make this work by consolidating pallets, reducing the actual number of pallets used.
                # This isn't handled by the algorithm, so - for the sake of fairness - any routes that have more pallets
                # than their vehicle's capacity will have their demand reduced.
                total_delivered = sum([stop.delivered for stop in route.stops])
                stop_ind = 0
                while total_delivered > route.vehicle.vehicle_type.capacity:
                    # Cycle through the stops, removing one demand from each at a time until total demand is acceptable
                    route.stops[stop_ind].delivered -= 1
                    route.stops[stop_ind].location.demand -= 1
                    total_delivered -= 1
                    stop_ind = (stop_ind + 1) % len(route.stops)
                if not routes.get(route.vehicle.vehicle_type.data_index):
                    routes[route.vehicle.vehicle_type.data_index] = []
                routes[route.vehicle.vehicle_type.data_index].append(route)
            # Check if the current route code matches anything already found
            # (just in case one vehicle's lines are spread out)
            route = find_route(code, routes)
            # If not, create a new route
            if not route:
                route = Route(code=code, vehicle=find_vehicle(horse=deliveries_sheet[f"J{row}"].value,
                                                              trailer=deliveries_sheet[f"K{row}"].value,
                                                              carried=deliveries_sheet[f"P{row}"].value,
                                                              vehicles=vehicles, vehicle_types=vehicle_types))
                route.vehicle.vehicle_type.vehicles_used += 1

        # Find the location
        location = find_location(deliveries_sheet[f"A{row}"].value, locations)

        # Read in the demands, defaulting to 0 if they are blank or have "C/S"
        try:
            dry_demand = int(deliveries_sheet[f"M{row}"].value)
        except ValueError:
            dry_demand = 0
        try:
            perish_demand = deliveries_sheet[f"N{row}"].value
        except ValueError:
            perish_demand = 0
        try:
            pick_by_line_demand = deliveries_sheet[f"O{row}"].value
        except ValueError:
            pick_by_line_demand = 0
        demand = dry_demand + perish_demand + pick_by_line_demand

        # Add the demand to the location's overall demand
        location.demand += demand
        # And set the amount delivered on this stop in the route
        route.stops.append(Stop(location, demand))

        row += 1

    return routes


def save_archive_routes(routes: Dict[int, List[Route]]):
    """Save the archive routes to the workbook in JSON format."""
    workbook = load_workbook("Model Data.xlsx")
    workbook["Archive Routes"]["A1"].value = json.dumps(routes)


def load_archive_routes() -> Dict[int, List[Route]]:
    """Load the archive routes from the workbook in JSON format."""
    workbook = load_workbook("Model Data.xlsx")
    return json.loads(workbook["Archive Routes"]["A1"].value)


def save_input_data(locations: List[PhysicalLocation], vehicle_types: List[VehicleType], anonymised: bool = False):
    """Saves the information for all locations and vehicle types for the day imported."""
    workbook = load_workbook("Model Data.xlsx")

    # Add the locations
    locations_sheet: Worksheet = workbook["Locations"]
    for index, location in enumerate(locations):
        if anonymised:
            locations_sheet[f"A{index + 2}"].value = location.anonymous_name
        else:
            locations_sheet[f"A{index + 2}"].value = location.name
            locations_sheet[f"B{index + 2}"].value = location.latitude
            locations_sheet[f"C{index + 2}"].value = location.longitude
        locations_sheet[f"D{index + 2}"].value = location.demand
        locations_sheet[f"E{index + 2}"].value = 0
        locations_sheet[f"F{index + 2}"].value = 24
        locations_sheet[f"G{index + 2}"].value = location.average_offload_time

    # Add the location distances and times
    # distances_sheet: Worksheet = workbook["Distances"]
    # times_sheet: Worksheet = workbook["Times"]
    # for row, distances_row in enumerate(distances):
    #     # Put location names at the start of each row
    #     if anonymised:
    #         distances_sheet.cell(row=row + 2, column=1, value=locations[row].anonymous_name)
    #         times_sheet.cell(row=row + 2, column=1, value=locations[row].anonymous_name)
    #     else:
    #         distances_sheet.cell(row=row + 2, column=1, value=locations[row].name)
    #         times_sheet.cell(row=row + 2, column=1, value=locations[row].name)
    #     for col, distance in enumerate(distances_row):
    #         if row == 0:
    #             # Put location names at the top of each column
    #             if anonymised:
    #                 distances_sheet.cell(row=1, column=col + 2, value=locations[col].anonymous_name)
    #                 times_sheet.cell(row=1, column=col + 2, value=locations[col].anonymous_name)
    #             else:
    #                 distances_sheet.cell(row=1, column=col + 2, value=locations[col].name)
    #                 times_sheet.cell(row=1, column=col + 2, value=locations[col].name)
    #         if row == col:
    #             # If on the same row and col, set the travel to be equal to going to the depot and back
    #             dist = distances[row][0] + distances[0][row]
    #             time = times[row][0] + times[0][row]
    #         else:
    #             dist = distances[row][col]
    #             time = times[row][col]
    #         # Apply the value to the sheet
    #         distances_sheet.cell(row=row + 2, column=col + 2, value=dist)
    #         times_sheet.cell(row=row + 2, column=col + 2, value=time)

    # Update the used vehicles counts
    vehicle_types_sheet: Worksheet = workbook["Vehicle Types"]
    for vehicle_type in vehicle_types:
        vehicle_types_sheet[f"F{vehicle_type.data_index + 2}"].value = vehicle_type.vehicles_used

    workbook.save("Model Data.xlsx")


def convert_archive(archive_file: str, anonymised: bool = False):
    """Reads in the demand, location, vehicle, and route data from the archive and then saves it"""
    locations = read_locations(filename="SPAR Locations and Schedule.xlsx")
    vehicle_types, vehicles = read_vehicles(types_file="Model Data.xlsx", types_sheet="Vehicle Types",
                                            vehicles_file="Spar Fleet.xlsx", vehicles_sheet="trucks")
    routes = read_archive(archive_file, locations, vehicles, vehicle_types)
    save_input_data(locations, vehicle_types, anonymised)
    save_archive_routes(routes)
    if anonymised:
        location_names = [location.anonymous_name for location in locations]
    else:
        location_names = [location.name for location in locations]


def import_data() -> Data:
    """Reads in the demand, location, and vehicle data to create input data for the algorithm.
    Also generates the routes from the archived data."""
    data_workbook = load_workbook(filename="Model Data.xlsx", read_only=True)

    # Read in the location information
    location_sheet = data_workbook["Locations"]
    locations: List[str] = []
    demand: List[float] = []
    window_start: List[float] = []
    window_end: List[float] = []
    average_unload_time: List[float] = []
    row = 2
    while location_sheet[f"A{row}"].value:
        locations.append(location_sheet[f"A{row}"].value)
        demand.append(location_sheet[f"D{row}"].value)
        window_start.append(location_sheet[f"E{row}"].value)
        window_end.append(location_sheet[f"F{row}"].value)
        average_unload_time.append(location_sheet[f"G{row}"].value)

        row += 1

    # Read in the vehicle data
    vehicle_sheet = data_workbook["Vehicle types"]
    vehicle_types: List[str] = []
    distance_cost: List[float] = []
    time_cost: List[float] = []
    pallet_capacity: List[int] = []
    available_vehicles: List[int] = []
    row = 2
    while vehicle_sheet[f"A{row}"].value:
        vehicle_types.append(vehicle_sheet[f"A{row}"].value)
        distance_cost.append(vehicle_sheet[f"B{row}"].value)
        time_cost.append(vehicle_sheet[f"C{row}"].value)
        pallet_capacity.append(vehicle_sheet[f"D{row}"].value)
        available_vehicles.append(vehicle_sheet[f"G{row}"].value)

        row += 1

    distances, times = read_input_data()

    archive_data = Data(locations=locations, demand=demand, window_start=window_start, window_end=window_end,
                        average_unload_time=average_unload_time, distances=distances, times=times,
                        vehicle_types=vehicle_types, distance_cost=distance_cost, time_cost=time_cost,
                        pallet_capacity=pallet_capacity, available_vehicles=available_vehicles)

    return archive_data
