from __future__ import annotations

from dataclasses import dataclass, field
from types import TracebackType
from typing import Callable, List, Optional


@dataclass
class FakeConnection:
    """Fake QBConnection that records sent requests and returns canned responses.

    If `responses` has fewer elements than `send` calls, the last response is
    reused. Responses may also be callables taking the request xml and returning
    the response xml (useful for dispatching per-request-type).
    """

    responses: List[object] = field(default_factory=list)
    sent_requests: List[str] = field(default_factory=list)

    def send(self, request_xml: str) -> str:
        self.sent_requests.append(request_xml)
        idx = min(len(self.sent_requests) - 1, len(self.responses) - 1)
        r = self.responses[idx]
        if callable(r):
            result = r(request_xml)
            assert isinstance(result, str)
            return result
        assert isinstance(r, str)
        return r

    def __enter__(self) -> "FakeConnection":
        return self

    def __exit__(
        self,
        exc_type: Optional[type],
        exc: Optional[BaseException],
        tb: Optional[TracebackType],
    ) -> None:
        return None


def make_factory(*responses: object) -> Callable[[], FakeConnection]:
    """Return a connection_factory that creates a FakeConnection with the given responses."""
    resp_list = list(responses)

    def factory() -> FakeConnection:
        return FakeConnection(responses=resp_list)

    return factory
